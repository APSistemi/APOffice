using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Windows.Forms;

namespace APOffice
{
    public partial class frmAnaArtImgWebSearch : Form
    {
        private string _artCod = "";
        private string _artDes = "";
        private List<string> _eanList = new List<string>();

        public string SelectedImagePath { get; private set; } = "";

        private class WebImageItem
        {
            public string Url { get; set; }
            public string SourceName { get; set; }
            public Image LoadedImage { get; set; }
            public Panel CardPanel { get; set; }
            public PictureBox CardPic { get; set; }
            public Label CardLabel { get; set; }
        }

        private List<WebImageItem> _foundImages = new List<WebImageItem>();
        private WebImageItem _selectedItem = null;
        private HashSet<string> _seenUrls = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        private int _activeSearchId = 0;

        public frmAnaArtImgWebSearch(string artCod, string artDes, List<string> eanList)
        {
            InitializeComponent();
            _artCod = artCod ?? "";
            _artDes = artDes ?? "";
            _eanList = eanList ?? new List<string>();
        }

        private void frmAnaArtImgWebSearch_Load(object sender, EventArgs e)
        {
            // Popola cmbQuery con i barcode e la descrizione
            cmbQuery.Items.Clear();

            foreach (string ean in _eanList)
            {
                if (!string.IsNullOrEmpty(ean) && !cmbQuery.Items.Contains(ean))
                    cmbQuery.Items.Add(ean);
            }

            if (!string.IsNullOrEmpty(_artDes) && !cmbQuery.Items.Contains(_artDes))
                cmbQuery.Items.Add(_artDes);

            if (_eanList.Count > 0 && !string.IsNullOrEmpty(_artDes))
            {
                string combo = _eanList[0] + " " + _artDes;
                if (!cmbQuery.Items.Contains(combo))
                    cmbQuery.Items.Add(combo);
            }

            if (cmbQuery.Items.Count > 0)
                cmbQuery.SelectedIndex = 0;
            else if (!string.IsNullOrEmpty(_artCod))
                cmbQuery.Text = _artCod;

            // Avvia automaticamente la ricerca iniziale
            EseguiRicerca();
        }

        private void btnCerca_Click(object sender, EventArgs e)
        {
            EseguiRicerca();
        }

        private void cmbQuery_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Return)
            {
                e.Handled = true;
                e.SuppressKeyPress = true;
                EseguiRicerca();
            }
        }

        private void EseguiRicerca()
        {
            string query = cmbQuery.Text.Trim();
            if (string.IsNullOrEmpty(query)) return;

            // Reset stato UI
            flowGallery.Controls.Clear();
            _foundImages.Clear();
            _seenUrls.Clear();
            _selectedItem = null;
            picBigPreview.Image = null;
            lblPreviewInfo.Text = "Ricerca in corso...\nLe immagini appariranno nella galleria.";
            btnAbbina.Enabled = false;
            btnApriBrowser.Enabled = false;
            progressBarSearch.Visible = true;
            btnCerca.Enabled = false;
            lblSearchStatus.Text = "⏳ Interrogazione motori di ricerca in corso...";
            lblSearchStatus.ForeColor = Color.FromArgb(37, 99, 235);
            lblStatusBottom.Text = "Ricerca in corso...";

            int currentSearchId = Interlocked.Increment(ref _activeSearchId);

            ThreadPool.QueueUserWorkItem(state =>
            {
                try
                {
                    // Force TLS 1.2 & TLS 1.1 support
                    ServicePointManager.SecurityProtocol = (SecurityProtocolType)3072 | (SecurityProtocolType)768 | SecurityProtocolType.Tls;

                    // 1. Ricerca tramite Barcode noti
                    List<string> queriesToSearch = new List<string>();
                    
                    // Aggiungi barcode
                    foreach (string ean in _eanList)
                    {
                        if (!string.IsNullOrEmpty(ean) && !queriesToSearch.Contains(ean))
                            queriesToSearch.Add(ean);
                    }

                    // Aggiungi query digitata dall'utente se diversa
                    if (!queriesToSearch.Contains(query))
                        queriesToSearch.Insert(0, query);

                    // Aggiungi descrizione articolo
                    if (!string.IsNullOrEmpty(_artDes) && !queriesToSearch.Contains(_artDes))
                        queriesToSearch.Add(_artDes);

                    foreach (string q in queriesToSearch)
                    {
                        if (currentSearchId != _activeSearchId || this.IsDisposed) return;

                        // Open Food Facts / Beauty / Pet se numerico
                        if (IsNumeric(q) && q.Length >= 8)
                        {
                            SearchOpenFoodFacts(q, currentSearchId);
                            SearchUpcItemDb(q, currentSearchId);
                        }

                        // Yahoo Images
                        SearchYahooImages(q, currentSearchId);

                        // Bing Images
                        SearchBingImages(q, currentSearchId);

                        // Google Images
                        SearchGoogleImages(q, currentSearchId);
                    }
                }
                catch { }
                finally
                {
                    if (!this.IsDisposed && this.IsHandleCreated)
                    {
                        this.Invoke(new Action(() =>
                        {
                            if (currentSearchId == _activeSearchId)
                            {
                                progressBarSearch.Visible = false;
                                btnCerca.Enabled = true;
                                lblSearchStatus.Text = "✅ Ricerca completata.";
                                lblSearchStatus.ForeColor = Color.FromArgb(22, 163, 74);
                                lblStatusBottom.Text = $"Trovate {_foundImages.Count} immagini online.";

                                if (_foundImages.Count == 0)
                                {
                                    Label lblEmpty = new Label();
                                    lblEmpty.Text = "⚠️ Nessuna immagine trovata per la query specificata.\nProva a cercare per descrizione del prodotto o per codice a barre.";
                                    lblEmpty.AutoSize = false;
                                    lblEmpty.Size = new Size(500, 80);
                                    lblEmpty.Font = new Font("Segoe UI", 10F, FontStyle.Regular);
                                    lblEmpty.ForeColor = Color.FromArgb(100, 116, 139);
                                    lblEmpty.TextAlign = ContentAlignment.MiddleCenter;
                                    flowGallery.Controls.Add(lblEmpty);
                                    lblPreviewInfo.Text = "Nessuna immagine trovata.";
                                }
                            }
                        }));
                    }
                }
            });
        }

        private bool IsNumeric(string s)
        {
            if (string.IsNullOrEmpty(s)) return false;
            foreach (char c in s)
            {
                if (!char.IsDigit(c)) return false;
            }
            return true;
        }

        #region Search Providers

        private void SearchOpenFoodFacts(string ean, int searchId)
        {
            string[] endpoints = new string[]
            {
                $"https://it.openfoodfacts.org/api/v0/product/{ean}.json",
                $"https://world.openfoodfacts.org/api/v0/product/{ean}.json",
                $"https://it.openbeautyfacts.org/api/v0/product/{ean}.json",
                $"https://world.openproductsfacts.org/api/v0/product/{ean}.json",
                $"https://world.openpetfoodfacts.org/api/v0/product/{ean}.json"
            };

            foreach (string endpoint in endpoints)
            {
                if (searchId != _activeSearchId || this.IsDisposed) return;
                try
                {
                    string json = FetchUrl(endpoint, "APOffice Retail Manager/1.0");
                    if (string.IsNullOrEmpty(json) || !json.Contains("\"product\":")) continue;

                    string[] keys = new string[] {
                        "image_front_url", "image_url", "image_small_url", "image_front_small_url",
                        "image_nutrition_url", "image_ingredients_url", "image_packaging_url"
                    };

                    foreach (string key in keys)
                    {
                        string val = ExtractJsonString(json, key);
                        if (!string.IsNullOrEmpty(val) && val.StartsWith("http", StringComparison.OrdinalIgnoreCase))
                        {
                            RegisterImage(val, "Open Food Facts", searchId);
                        }
                    }

                    // Estrai immagini ad alta risoluzione selezionate
                    MatchCollection matches = Regex.Matches(json, @"https://images\.openfoodfacts\.org/[^""\s,]+", RegexOptions.IgnoreCase);
                    foreach (Match m in matches)
                    {
                        RegisterImage(m.Value, "Open Food Facts", searchId);
                    }
                }
                catch { }
            }
        }

        private void SearchUpcItemDb(string ean, int searchId)
        {
            try
            {
                string url = "https://api.upcitemdb.com/prod/trial/lookup?upc=" + ean;
                string json = FetchUrl(url, "Mozilla/5.0 (Windows NT 10.0; Win64; x64)");
                if (string.IsNullOrEmpty(json)) return;

                MatchCollection matches = Regex.Matches(json, @"https?://[^""\s,]+\.(?:jpg|jpeg|png|webp)", RegexOptions.IgnoreCase);
                foreach (Match m in matches)
                {
                    RegisterImage(m.Value, "UPCItemDB", searchId);
                }
            }
            catch { }
        }

        private void SearchYahooImages(string query, int searchId)
        {
            try
            {
                string url = "https://it.images.search.yahoo.com/search/images?p=" + Uri.EscapeDataString(query);
                string html = FetchUrl(url, "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36");
                if (string.IsNullOrEmpty(html)) return;

                // Match imgurl
                MatchCollection mImgUrls = Regex.Matches(html, @"imgurl=(https?%3A%2F%2F[^&""]+)", RegexOptions.IgnoreCase);
                foreach (Match m in mImgUrls)
                {
                    string decoded = Uri.UnescapeDataString(m.Groups[1].Value);
                    if (decoded.StartsWith("http", StringComparison.OrdinalIgnoreCase))
                        RegisterImage(decoded, "Yahoo Images", searchId);
                }

                // Match data-src o src
                MatchCollection mDataSrc = Regex.Matches(html, @"data-src=""(https://[^""]+\.(?:jpg|jpeg|png|webp)[^""]*)""", RegexOptions.IgnoreCase);
                foreach (Match m in mDataSrc)
                {
                    string decoded = WebUtility.HtmlDecode(m.Groups[1].Value);
                    RegisterImage(decoded, "Yahoo Images", searchId);
                }

                MatchCollection mBingTh = Regex.Matches(html, @"<img[^>]+src=""(https://tse\d\.mm\.bing\.net/th\?[^""]+)""", RegexOptions.IgnoreCase);
                foreach (Match m in mBingTh)
                {
                    string decoded = WebUtility.HtmlDecode(m.Groups[1].Value);
                    RegisterImage(decoded, "Yahoo / Bing", searchId);
                }
            }
            catch { }
        }

        private void SearchBingImages(string query, int searchId)
        {
            try
            {
                string url = "https://www.bing.com/images/search?q=" + Uri.EscapeDataString(query);
                string html = FetchUrl(url, "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36");
                if (string.IsNullOrEmpty(html)) return;

                MatchCollection mUrls = Regex.Matches(html, @"murl&quot;:&quot;(https?://[^&""]+)&quot;", RegexOptions.IgnoreCase);
                foreach (Match m in mUrls)
                {
                    string decoded = WebUtility.HtmlDecode(m.Groups[1].Value);
                    RegisterImage(decoded, "Bing Images", searchId);
                }

                MatchCollection mTh = Regex.Matches(html, @"src=""(https://tse\d\.mm\.bing\.net/th\?[^""]+)""", RegexOptions.IgnoreCase);
                foreach (Match m in mTh)
                {
                    string decoded = WebUtility.HtmlDecode(m.Groups[1].Value);
                    RegisterImage(decoded, "Bing Images", searchId);
                }
            }
            catch { }
        }

        private void SearchGoogleImages(string query, int searchId)
        {
            try
            {
                string url = "https://www.google.it/search?tbm=isch&q=" + Uri.EscapeDataString(query) + "&hl=it&gbv=1";
                string html = FetchUrl(url, "Mozilla/5.0 (Windows NT 10.0; Win64; x64)");
                if (string.IsNullOrEmpty(html)) return;

                MatchCollection mTh = Regex.Matches(html, @"src=""(https://encrypted-tbn0\.gstatic\.com/images\?[^""]+)""", RegexOptions.IgnoreCase);
                foreach (Match m in mTh)
                {
                    string decoded = WebUtility.HtmlDecode(m.Groups[1].Value);
                    RegisterImage(decoded, "Google Images", searchId);
                }
            }
            catch { }
        }

        #endregion

        #region Helpers & UI Card Rendering

        private string FetchUrl(string url, string userAgent)
        {
            try
            {
                HttpWebRequest req = (HttpWebRequest)WebRequest.Create(url);
                req.UserAgent = userAgent;
                req.Headers.Add("Accept-Language", "it-IT,it;q=0.9,en-US;q=0.8,en;q=0.7");
                req.AutomaticDecompression = DecompressionMethods.GZip | DecompressionMethods.Deflate;
                req.Timeout = 6000;
                req.ReadWriteTimeout = 6000;

                using (HttpWebResponse resp = (HttpWebResponse)req.GetResponse())
                using (Stream stream = resp.GetResponseStream())
                using (StreamReader sr = new StreamReader(stream, Encoding.UTF8))
                {
                    return sr.ReadToEnd();
                }
            }
            catch
            {
                return "";
            }
        }

        private string ExtractJsonString(string json, string key)
        {
            if (string.IsNullOrEmpty(json)) return "";
            try
            {
                string pattern = "\"" + Regex.Escape(key) + "\":\\s*\"((?:\\\\\"|[^\"])*)\"";
                Match m = Regex.Match(json, pattern);
                if (m.Success)
                    return m.Groups[1].Value.Replace("\\/", "/").Replace("\\\"", "\"").Trim();
            }
            catch { }
            return "";
        }

        private void RegisterImage(string url, string sourceName, int searchId)
        {
            if (string.IsNullOrEmpty(url) || !url.StartsWith("http", StringComparison.OrdinalIgnoreCase)) return;

            // Normalize URL for deduplication
            string normUrl = url.Split('?')[0].ToLowerInvariant();
            if (_seenUrls.Contains(normUrl)) return;
            _seenUrls.Add(normUrl);

            WebImageItem item = new WebImageItem
            {
                Url = url,
                SourceName = sourceName
            };

            if (this.IsDisposed || !this.IsHandleCreated) return;

            this.Invoke(new Action(() =>
            {
                if (searchId != _activeSearchId) return;

                _foundImages.Add(item);
                AddCardToGallery(item, searchId);
                lblStatusBottom.Text = $"Trovate {_foundImages.Count} immagini online...";
            }));
        }

        private void AddCardToGallery(WebImageItem item, int searchId)
        {
            Panel card = new Panel();
            card.Size = new Size(146, 178);
            card.BackColor = Color.White;
            card.BorderStyle = BorderStyle.FixedSingle;
            card.Margin = new Padding(6);
            card.Cursor = Cursors.Hand;

            PictureBox pic = new PictureBox();
            pic.Size = new Size(138, 130);
            pic.Location = new Point(3, 3);
            pic.SizeMode = PictureBoxSizeMode.Zoom;
            pic.BackColor = Color.FromArgb(248, 250, 252);
            pic.Cursor = Cursors.Hand;

            Label lbl = new Label();
            lbl.Location = new Point(3, 136);
            lbl.Size = new Size(138, 36);
            lbl.Font = new Font("Segoe UI", 7.5F, FontStyle.Regular);
            lbl.ForeColor = Color.FromArgb(100, 116, 139);
            lbl.TextAlign = ContentAlignment.MiddleCenter;
            lbl.Text = "Caricamento...";
            lbl.Cursor = Cursors.Hand;

            card.Controls.Add(pic);
            card.Controls.Add(lbl);

            item.CardPanel = card;
            item.CardPic = pic;
            item.CardLabel = lbl;

            // Click & Selection handlers
            EventHandler onSelect = (s, e) => SelectImageItem(item);
            EventHandler onDblClick = (s, e) =>
            {
                SelectImageItem(item);
                ConfirmSelectionAndSave();
            };

            card.Click += onSelect;
            pic.Click += onSelect;
            lbl.Click += onSelect;

            card.DoubleClick += onDblClick;
            pic.DoubleClick += onDblClick;
            lbl.DoubleClick += onDblClick;

            // Hover effects
            card.MouseEnter += (s, e) => { if (_selectedItem != item) card.BackColor = Color.FromArgb(239, 246, 255); };
            card.MouseLeave += (s, e) => { if (_selectedItem != item) card.BackColor = Color.White; };
            pic.MouseEnter += (s, e) => { if (_selectedItem != item) card.BackColor = Color.FromArgb(239, 246, 255); };
            lbl.MouseEnter += (s, e) => { if (_selectedItem != item) card.BackColor = Color.FromArgb(239, 246, 255); };

            flowGallery.Controls.Add(card);

            // Asynchronous thumbnail download
            ThreadPool.QueueUserWorkItem(state =>
            {
                Image downloaded = DownloadImageFromUrl(item.Url);
                if (downloaded == null || this.IsDisposed || !this.IsHandleCreated) return;

                this.Invoke(new Action(() =>
                {
                    if (searchId != _activeSearchId) return;

                    item.LoadedImage = downloaded;
                    pic.Image = downloaded;
                    lbl.Text = $"{downloaded.Width}x{downloaded.Height}\n{item.SourceName}";

                    // Se è la prima immagine caricata, selezionala come anteprima predefinita
                    if (_selectedItem == null)
                    {
                        SelectImageItem(item);
                    }
                }));
            });
        }

        private Image DownloadImageFromUrl(string url)
        {
            try
            {
                ServicePointManager.SecurityProtocol = (SecurityProtocolType)3072 | (SecurityProtocolType)768 | SecurityProtocolType.Tls;
                HttpWebRequest req = (HttpWebRequest)WebRequest.Create(url);
                req.UserAgent = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36";
                req.Headers.Add("Accept-Language", "it-IT,it;q=0.9,en-US;q=0.8,en;q=0.7");
                req.AutomaticDecompression = DecompressionMethods.GZip | DecompressionMethods.Deflate;
                req.Timeout = 7000;
                req.ReadWriteTimeout = 7000;

                using (HttpWebResponse resp = (HttpWebResponse)req.GetResponse())
                using (Stream stream = resp.GetResponseStream())
                using (MemoryStream ms = new MemoryStream())
                {
                    stream.CopyTo(ms);
                    ms.Position = 0;
                    using (Image rawImg = Image.FromStream(ms))
                    {
                        return new Bitmap(rawImg);
                    }
                }
            }
            catch
            {
                return null;
            }
        }

        private void SelectImageItem(WebImageItem item)
        {
            if (item == null) return;

            // Deseleziona precedente
            if (_selectedItem != null && _selectedItem.CardPanel != null)
            {
                _selectedItem.CardPanel.BackColor = Color.White;
                _selectedItem.CardPanel.Padding = new Padding(0);
            }

            _selectedItem = item;

            // Evidenzia card selezionata
            if (item.CardPanel != null)
            {
                item.CardPanel.BackColor = Color.FromArgb(219, 234, 254);
            }

            // Aggiorna pannello anteprima
            if (item.LoadedImage != null)
            {
                picBigPreview.Image = item.LoadedImage;
                lblPreviewInfo.Text = $"Risoluzione: {item.LoadedImage.Width} x {item.LoadedImage.Height} px\n" +
                                       $"Sorgente: {item.SourceName}\n" +
                                       $"URL: {(item.Url.Length > 60 ? item.Url.Substring(0, 57) + "..." : item.Url)}";
                btnAbbina.Enabled = true;
                btnApriBrowser.Enabled = true;
            }
            else
            {
                lblPreviewInfo.Text = $"Caricamento in corso...\nSorgente: {item.SourceName}\nURL: {item.Url}";
                btnAbbina.Enabled = true;
                btnApriBrowser.Enabled = true;
            }
        }

        private void ConfirmSelectionAndSave()
        {
            if (_selectedItem == null || string.IsNullOrEmpty(_selectedItem.Url))
            {
                MessageBox.Show("Seleziona prima un'immagine dalla galleria.", "Nessuna selezione", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            if (string.IsNullOrEmpty(_artCod) || _artCod == "NEW" || _artCod == "NEWA")
            {
                MessageBox.Show("Codice articolo non valido.", "Attenzione", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            try
            {
                string targetDir = @"C:\ApProject\Temp\Img";
                if (!Directory.Exists(targetDir))
                {
                    Directory.CreateDirectory(targetDir);
                }

                string destPath = Path.Combine(targetDir, _artCod + ".jpg");

                // Se l'immagine è già in memoria come Bitmap, la salviamo in alta qualità
                if (_selectedItem.LoadedImage != null)
                {
                    using (Bitmap bmp = new Bitmap(_selectedItem.LoadedImage))
                    {
                        bmp.Save(destPath, System.Drawing.Imaging.ImageFormat.Jpeg);
                    }
                }
                else
                {
                    // Altrimenti scarica direttamente da stream URL
                    ServicePointManager.SecurityProtocol = (SecurityProtocolType)3072 | (SecurityProtocolType)768 | SecurityProtocolType.Tls;
                    using (WebClient wc = new WebClient())
                    {
                        wc.Headers.Add("User-Agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36");
                        wc.DownloadFile(_selectedItem.Url, destPath);
                    }
                }

                if (File.Exists(destPath))
                {
                    SelectedImagePath = destPath;
                    this.DialogResult = DialogResult.OK;
                    this.Close();
                }
                else
                {
                    MessageBox.Show("Impossibile salvare il file immagine nella cartella destinazione.", "Errore Salvataggio", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Errore durante il download e l'abbinamento dell'immagine:\n" + ex.Message, "Errore", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btnAbbina_Click(object sender, EventArgs e)
        {
            ConfirmSelectionAndSave();
        }

        private void btnApriBrowser_Click(object sender, EventArgs e)
        {
            if (_selectedItem != null && !string.IsNullOrEmpty(_selectedItem.Url))
            {
                try
                {
                    System.Diagnostics.Process.Start(_selectedItem.Url);
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Impossibile aprire il browser:\n" + ex.Message, "Errore", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private void btnChiudi_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.Cancel;
            this.Close();
        }

        #endregion
    }
}
