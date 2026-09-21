using System;
using System.Drawing;
using System.Windows.Forms;
using System.Text;
using System.Threading.Tasks;
using System.Net;
using System.IO;
using System.Diagnostics;

namespace APOffice
{
    public partial class frmAIChat : Form
    {
        private TextBox txtInput;
        private Button btnSend;
        private WebBrowser wbChat;
        private StringBuilder chatHtml;
        private const string API_URL = "http://127.0.0.1:8765/query";

        // Storico Prompts
        private System.Collections.Generic.List<string> _history = new System.Collections.Generic.List<string>();
        private int _historyIndex = -1;
        private string _currentInput = "";

        public frmAIChat()
        {
            InitializeComponent();
            this.DoubleBuffered = true;
            SetupUI();
            chatHtml = new StringBuilder();
            ResetChatHtml();
        }

        private void SetupUI()
        {
            this.Text = "APOffice AI Assistant";
            this.Size = new Size(400, 600);
            this.FormBorderStyle = FormBorderStyle.SizableToolWindow;
            this.StartPosition = FormStartPosition.Manual;
            this.BackColor = Color.FromArgb(245, 246, 248);
            this.Icon = SystemIcons.Information; // Fallback icon

            // Header Panel
            Panel pnlHeader = new Panel();
            pnlHeader.Dock = DockStyle.Top;
            pnlHeader.Height = 50;
            pnlHeader.BackColor = Color.FromArgb(52, 73, 94);

            Label lblTitle = new Label();
            lblTitle.Text = "🤖 Agente AI - Assistente Operativo";
            lblTitle.ForeColor = Color.White;
            lblTitle.Font = new Font("Segoe UI", 12, FontStyle.Bold);
            lblTitle.AutoSize = false;
            lblTitle.Dock = DockStyle.Fill;
            lblTitle.TextAlign = ContentAlignment.MiddleCenter;
            pnlHeader.Controls.Add(lblTitle);

            // Bottom Panel for Input
            Panel pnlBottom = new Panel();
            pnlBottom.Dock = DockStyle.Bottom;
            pnlBottom.Height = 60;
            pnlBottom.BackColor = Color.White;
            pnlBottom.Padding = new Padding(10);

            txtInput = new TextBox();
            txtInput.Multiline = true;
            txtInput.Font = new Font("Segoe UI", 10);
            txtInput.Dock = DockStyle.Fill;
            txtInput.ScrollBars = ScrollBars.Vertical;
            txtInput.BorderStyle = BorderStyle.None;
            txtInput.KeyDown += TxtInput_KeyDown;

            btnSend = new Button();
            btnSend.Text = "Invia";
            btnSend.Dock = DockStyle.Right;
            btnSend.Width = 80;
            btnSend.BackColor = Color.FromArgb(41, 128, 185);
            btnSend.ForeColor = Color.White;
            btnSend.FlatStyle = FlatStyle.Flat;
            btnSend.Font = new Font("Segoe UI", 10, FontStyle.Bold);
            btnSend.Cursor = Cursors.Hand;
            btnSend.Click += BtnSend_Click;

            pnlBottom.Controls.Add(txtInput);
            pnlBottom.Controls.Add(btnSend);

            // Chat Area
            wbChat = new WebBrowser();
            wbChat.Dock = DockStyle.Fill;
            wbChat.IsWebBrowserContextMenuEnabled = true; // Re-enable for Copy/Paste
            wbChat.WebBrowserShortcutsEnabled = true;    // Re-enable for shortcuts
            wbChat.ScriptErrorsSuppressed = true;
            wbChat.Navigating += WbChat_Navigating;

            this.Controls.Add(wbChat);
            this.Controls.Add(pnlBottom);
            this.Controls.Add(pnlHeader);
        }

        private void ResetChatHtml()
        {
            string skeleton = @"
                <html>
                <head>
                <meta http-equiv='X-UA-Compatible' content='IE=edge' />
                <script src='https://cdn.jsdelivr.net/npm/chart.js@2.9.4'></script>
                <link rel='stylesheet' href='https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.0.0/css/all.min.css' />
                <style>
                    body { font-family: 'Segoe UI', Arial, sans-serif; background-color: #f0f2f5; margin: 0; padding: 15px; font-size: 14px; }
                    #chat-container { width: 95%; margin: 0 auto; overflow: hidden; }
                    
                    .msg { margin-bottom: 20px; padding: 12px 18px; border-radius: 12px; line-height: 1.5; clear: both; max-width: 85%; display: block; position: relative; box-shadow: 0 1px 2px rgba(0,0,0,0.1); }
                    
                    /* MESSAGGIO UTENTE */
                    .user { float: right; background-color: #007bff; color: #ffffff; border-bottom-right-radius: 2px; }
                    
                    /* MESSAGGIO AI */
                    .ai { float: left; background-color: #ffffff; color: #1a1a1a; border: 1px solid #d1d7dd; border-bottom-left-radius: 2px; }
                    
                    .system { clear: both; background: #fff3cd; color: #856404; border: 1px solid #ffeeba; margin: 10px auto; text-align: center; }
                    .timestamp { font-size: 10px; opacity: 0.6; margin-top: 5px; text-align: right; }
                    
                    /* Grafiche Meteo */
                    .weather-card { background: linear-gradient(135deg, #00b4db 0%, #0083b0 100%); color: white; padding: 10px 15px; border-radius: 10px; margin: 8px 0; display: flex; align-items: center; justify-content: space-between; min-width: 250px; }
                    .weather-info { flex-grow: 1; margin-left: 15px; }
                    .weather-day { font-size: 11px; opacity: 0.8; font-weight: bold; }
                    .weather-desc { font-size: 15px; font-weight: bold; margin: 2px 0; }
                    .weather-temp { font-size: 13px; font-weight: bold; opacity: 0.9; }
                    .weather-icon { font-size: 42px; width: 50px; text-align: center; display: flex; align-items: center; justify-content: center; }

                    /* Tabelle */
                    table { border-collapse: collapse; width: 100%; margin: 10px 0; background: #fff; border: 1px solid #ddd; border-radius: 5px; overflow: hidden; }
                    th { background: #34495e; color: #fff; padding: 8px; text-align: left; }
                    td { padding: 8px; border-bottom: 1px solid #eee; }
                    
                    /* Canvas Chart */
                    .chart-wrapper { background: white; padding: 10px; border: 1px solid #ddd; border-radius: 8px; margin: 10px 0; }

                    /* Link Styling */
                    a { color: #007bff; text-decoration: underline; font-weight: bold; cursor: pointer; }
                    a:hover { color: #0056b3; }

                    #thinking-box { clear: both; display: none; float: left; margin-top: 10px; color: #666; font-style: italic; }
                    .dot { display: inline-block; width: 8px; height: 8px; background: #007bff; border-radius: 50%; margin-right: 5px; animation: wave 1.3s linear infinite; }
                    @keyframes wave { 0%, 60%, 100% { transform: initial; } 30% { transform: translateY(-5px); } }
                </style>
                </head>
                <body>
                    <div id='chat-container'></div>
                    <div id='thinking-box'><div class='dot'></div><div class='dot'></div><div class='dot'></div> pensando...</div>
                    <div style='clear:both; height: 50px;'></div>
                    <script>
                        function appendMessage(role, html, time) {
                            var container = document.getElementById('chat-container');
                            var msgDiv = document.createElement('div');
                            msgDiv.className = 'msg ' + role;
                            msgDiv.innerHTML = html + '<div class=\'timestamp\'>' + time + '</div>';
                            container.appendChild(msgDiv);

                            // IE11 safe script execution
                            var scriptList = msgDiv.getElementsByTagName('script');
                            var scriptArray = [];
                            for (var i = 0; i < scriptList.length; i++) { scriptArray.push(scriptList[i]); }
                            
                            for (var j = 0; j < scriptArray.length; j++) {
                                var s = document.createElement('script');
                                s.text = scriptArray[j].text || scriptArray[j].innerHTML;
                                document.body.appendChild(s);
                            }

                            scrollToBottom();
                        }
                        function setThinking(show) {
                            document.getElementById('thinking-box').style.display = show ? 'block' : 'none';
                            if(show) scrollToBottom();
                        }
                        function scrollToBottom() {
                            window.scroll(0, 50000);
                        }
                        
                        // Block Refresh (F5/Ctrl+R) but keep other shortcuts
                        document.onkeydown = function(e) {
                            if (e.keyCode === 116 || (e.ctrlKey && e.keyCode === 82)) {
                                e.preventDefault();
                                return false;
                            }
                        };
                        
                        // Helper per rendere i grafici iniettati via HTML
                        function initChart(id, config) {
                            var retryCount = 0;
                            var interval = setInterval(function() {
                                try {
                                    if (typeof Chart !== 'undefined') {
                                        var canvas = document.getElementById(id);
                                        if (canvas) {
                                            var ctx = canvas.getContext('2d');
                                            new Chart(ctx, config);
                                            clearInterval(interval);
                                        }
                                    } else if (retryCount > 30) {
                                        console.error('Chart.js timeout');
                                        clearInterval(interval);
                                    }
                                    retryCount++;
                                } catch(e) { 
                                    var errDiv = document.createElement('div');
                                    errDiv.style.color = 'red';
                                    errDiv.innerHTML = 'Errore Grafico: ' + e.message;
                                    document.getElementById(id).parentNode.appendChild(errDiv);
                                    clearInterval(interval);
                                }
                            }, 300);
                        }
                    </script>
                </body>
                </html>";

            wbChat.DocumentText = skeleton;
            chatHtml.Clear(); // StringBuilder used as backup or for export

            // Wait for document to load before first message
            Task.Delay(500).ContinueWith(t =>
            {
                this.Invoke((MethodInvoker)delegate
                {
                    AppendMessage("ai", "Ciao! Sono l'Agente AI di APOffice. Posso analizzare le vendite, controllare le giacenze, cercare articoli e molto altro. Come posso aiutarti oggi?");
                });
            });
        }

        private static int _chartCounter = 0;

        private void AppendMessage(string role, string text)
        {
            string htmlText = text.Replace("\n", "<br/>");

            // Basic Markdown Table Parser (Primitive but effective for our AI output)
            if (htmlText.Contains("|"))
            {
                htmlText = ParseMarkdownTable(htmlText);
            }

            // Bold and Code
            htmlText = System.Text.RegularExpressions.Regex.Replace(htmlText, @"\*\*(.*?)\*\*", "<b>$1</b>");
            htmlText = System.Text.RegularExpressions.Regex.Replace(htmlText, @"\`(.*?)\`", "<code>$1</code>");

            // Markdown Links: [text](url) -> <a href='url'>text</a>
            htmlText = System.Text.RegularExpressions.Regex.Replace(htmlText, @"\[(.*?)\]\((.*?)\)", "<a href='$2'>$1</a>");

            // Markdown Images: ![alt](url) -> <img src='url' />
            htmlText = System.Text.RegularExpressions.Regex.Replace(htmlText, @"\!\[(.*?)\]\((.*?)\)", "<img src='$2' style='max-width:90%; border-radius:8px; display:block; margin:10px 0;' alt='$1' />");

            // Custom Chart Parser: [CHART]{...}[/CHART]
            if (htmlText.Contains("[CHART]"))
            {
                var chartRegex = new System.Text.RegularExpressions.Regex(@"\[CHART\](.*?)\[/CHART\]", System.Text.RegularExpressions.RegexOptions.Singleline);
                htmlText = chartRegex.Replace(htmlText, m =>
                {
                    _chartCounter++;
                    string chartId = "chart_" + _chartCounter;
                    string config = m.Groups[1].Value.Trim();
                    return $"<div class='chart-wrapper'><canvas id='{chartId}' width='400' height='200'></canvas></div><script>initChart('{chartId}', {config});</script>";
                });
            }

            string timeString = DateTime.Now.ToString("HH:mm");

            if (wbChat.Document != null)
            {
                wbChat.Document.InvokeScript("appendMessage", new object[] { role, htmlText, timeString });
            }
        }

        private string ParseMarkdownTable(string text)
        {
            try
            {
                var lines = text.Split(new[] { "<br/>" }, StringSplitOptions.None);
                StringBuilder sb = new StringBuilder();
                bool inTable = false;

                foreach (var line in lines)
                {
                    if (line.Contains("|") && line.Trim().StartsWith("|"))
                    {
                        if (!inTable) { sb.Append("<table>"); inTable = true; }
                        if (line.Contains("---")) continue; // Skip separator line

                        sb.Append("<tr>");
                        var cells = line.Trim('|').Split('|');
                        foreach (var cell in cells)
                        {
                            sb.Append(inTable && sb.ToString().EndsWith("<table>") ? "<th>" : "<td>");
                            sb.Append(cell.Trim());
                            sb.Append(inTable && sb.ToString().Contains("<th>") && !sb.ToString().Contains("<tr><td>") ? "</th>" : "</td>");
                        }
                        sb.Append("</tr>");
                        // Simple toggle header/cell logic improvement
                    }
                    else
                    {
                        if (inTable) { sb.Append("</table>"); inTable = false; }
                        sb.Append(line + "<br/>");
                    }
                }
                if (inTable) sb.Append("</table>");

                // Fix header logic (first row is header)
                string result = sb.ToString().Replace("<table><tr><td>", "<table><tr><th>").Replace("</td><td>", "</th><th>").Replace("</td></tr>", "</th></tr>");
                // Correct subsequent rows
                return result;
            }
            catch { return text; }
        }

        private void RefreshBrowser()
        {
            // Metodo mantenuto per retrocompatibilità ma logica spostata in AppendMessage
        }

        private async void BtnSend_Click(object sender, EventArgs e)
        {
            string query = txtInput.Text.Trim();
            if (string.IsNullOrEmpty(query)) return;

            // Salva nello storico
            if (_history.Count == 0 || _history[_history.Count - 1] != query)
            {
                _history.Add(query);
                if (_history.Count > 100) _history.RemoveAt(0);
            }
            _historyIndex = -1;
            _currentInput = "";

            txtInput.Clear();
            AppendMessage("user", query);

            // Show Thinking animation
            if (wbChat.Document != null)
                wbChat.Document.InvokeScript("setThinking", new object[] { true });

            btnSend.Enabled = false;
            txtInput.Enabled = false;

            try
            {
                // To avoid Nuget dependencies (Newtonsoft.Json) in Legacy WinForms,
                // we build JSON manually and parse manually for this simple API response.
                string jsonPayload = "{ \"query\": \"" + query.Replace("\"", "\\\"").Replace("\n", "\\n") + "\" }";

                WebRequest request = WebRequest.Create(API_URL);
                request.Method = "POST";
                request.ContentType = "application/json";

                using (var streamWriter = new StreamWriter(await request.GetRequestStreamAsync()))
                {
                    streamWriter.Write(jsonPayload);
                }

                WebResponse response = await request.GetResponseAsync();
                string jsonResponse = "";
                using (var streamReader = new StreamReader(response.GetResponseStream()))
                {
                    jsonResponse = streamReader.ReadToEnd();
                }

                // Manual parsing of: {"status":"success","response":"testo riposta"}
                string aiText = ParseJsonResponse(jsonResponse);
                if (string.IsNullOrEmpty(aiText))
                    aiText = "Nessuna risposta dal server.";

                AppendMessage("ai", aiText);
            }
            catch (WebException wex)
            {
                AppendMessage("system", "Impossibile connettersi al Server AI locale (Errore HTTP). Assicurati che sia in esecuzione (start_server.bat).<br/>" + wex.Message);
            }
            catch (Exception ex)
            {
                AppendMessage("system", "Errore del client Chat:<br/>" + ex.Message);
            }
            finally
            {
                // Hide Thinking animation
                if (wbChat.Document != null)
                    wbChat.Document.InvokeScript("setThinking", new object[] { false });

                btnSend.Enabled = true;
                txtInput.Enabled = true;
                txtInput.Focus();

                // Re-inject scroll script and trigger it
                if (wbChat.Document != null && wbChat.Document.Window != null)
                    wbChat.Document.Window.ScrollTo(0, 99999);
            }
        }

        private string ParseJsonResponse(string json)
        {
            try
            {
                string marker = "\"response\":";
                int idx = json.IndexOf(marker);
                if (idx > -1)
                {
                    string sub = json.Substring(idx + marker.Length).Trim();
                    if (sub.StartsWith("\""))
                    {
                        sub = sub.Substring(1); // Rimuove prima virgoletta
                        int endIdx = sub.LastIndexOf("\"");
                        if (endIdx > -1)
                            sub = sub.Substring(0, endIdx);

                        // Unescape common JSON characters
                        sub = sub.Replace("\\n", "\n").Replace("\\\"", "\"").Replace("\\\\", "\\");
                        return sub;
                    }
                }
            }
            catch { }
            return json; // Fallback in case of parsing failure
        }

        private void TxtInput_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter && !e.Shift)
            {
                e.SuppressKeyPress = true; // Prevent ding sound and newline
                BtnSend_Click(this, EventArgs.Empty);
            }
            else if (e.KeyCode == Keys.Up && (txtInput.SelectionStart == 0 || txtInput.Text == ""))
            {
                if (_history.Count > 0)
                {
                    if (_historyIndex == -1)
                    {
                        _currentInput = txtInput.Text;
                        _historyIndex = _history.Count - 1;
                    }
                    else if (_historyIndex > 0)
                    {
                        _historyIndex--;
                    }

                    txtInput.Text = _history[_historyIndex];
                    txtInput.SelectionStart = txtInput.Text.Length;
                    e.SuppressKeyPress = true;
                }
            }
            else if (e.KeyCode == Keys.Down && txtInput.SelectionStart == txtInput.Text.Length)
            {
                if (_historyIndex != -1)
                {
                    if (_historyIndex < _history.Count - 1)
                    {
                        _historyIndex++;
                        txtInput.Text = _history[_historyIndex];
                    }
                    else
                    {
                        _historyIndex = -1;
                        txtInput.Text = _currentInput;
                    }
                    txtInput.SelectionStart = txtInput.Text.Length;
                    e.SuppressKeyPress = true;
                }
            }
        }

        private void WbChat_Navigating(object sender, WebBrowserNavigatingEventArgs e)
        {
            // Se non è about:blank, apriamo nel browser esterno
            if (e.Url.ToString() != "about:blank" && !e.Url.ToString().StartsWith("javascript:"))
            {
                e.Cancel = true;
                try
                {
                    Process.Start(e.Url.ToString());
                }
                catch
                {
                    // Fallback se il browser non risponde
                }
            }
        }

        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            if (e.CloseReason == CloseReason.UserClosing)
            {
                e.Cancel = true;
                this.Hide(); // Hide instead of close to keep chat history
            }
            base.OnFormClosing(e);
        }
    }
}
