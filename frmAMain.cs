using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using System.Diagnostics;
//using System.Runtime.InteropServices;

namespace APOffice
{
    public partial class frmAMain : Form
    {
        //[DllImport("pagamento.dll")]
        ////private static extern void pagamento();
        //public static extern void AnnullaPIN();

        clsDefine _clsDef = new clsDefine();
        clsFuncs _clsFun = new clsFuncs();
        clsQuery _clsQry = new clsQuery();

        [System.Runtime.InteropServices.DllImport("user32.dll")]
        [return: System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.Bool)]
        private static extern bool SetForegroundWindow(IntPtr hWnd);

        private string _strConSql = "";
        private List<Button> _shortcutButtons = new List<Button>();

        public frmAMain()
        {
            InitializeComponent();
            new clsGesGraph().SetGraph(this, 0);
        }

        private void frmAMain_Load(object sender, EventArgs e)
        {
            _strConSql = _clsFun.ConSql("");

            // Caricamento Ragione Sociale (già verificata in Splash, qui la recuperiamo per il titolo)
            string s = _clsFun.FileIni("R", clsDefine.enuIni.Ini09CodiceAzienda, "");
            if (s == "") s = "001";
            DataTable t = new clsQuery().ConfAzienda(s);
            string appVer = System.Reflection.Assembly.GetExecutingAssembly().GetName().Version.ToString();
            this.Text = "AP OFFICE v" + appVer;

            if (t.Rows.Count > 0)
            {
                this.Text = "AP OFFICE v" + appVer + " - " + (string)t.Rows[0]["cnf_rag"];
                if ((string)t.Rows[0]["cnf_rag"] == "AP Sistemi SRL")
                    this.BackColor = Color.Brown;

                // Controllo credenziali se abilitate
                s = _clsFun.FileIni("R", clsDefine.enuIni.IniCrenziali, "");
                if (s != "N")
                {
                    frmUser frmUser = new frmUser();
                    frmUser._strConSql = _strConSql;
                    if (frmUser.ShowDialog() != DialogResult.OK && !frmUser._bolSta)
                    {
                        this.Close();
                        return;
                    }
                }

                Controlli();
                CtrlDivulgazioni();
                CtrlChiusuraBat();
                CtrlOfferte();

                // Menu Utility Etichette
                ToolStripMenuItem mnuUty = new ToolStripMenuItem("Utilità Etichette");
                mnuUty.Image = clsUiIcons.GetIcon("labels", 16);
                ToolStripMenuItem mnuFmt = new ToolStripMenuItem("Formati Etichette Personalizzati");
                mnuFmt.Image = clsUiIcons.GetIcon("barcode", 16);
                mnuFmt.Click += (s2, ev2) => new frmUtyEtiFormati().ShowDialog();
                mnuUty.DropDownItems.Add(mnuFmt);
                this.menuStrip1.Items.Add(mnuUty);

                // Menu Agente AI se abilitato in ini
                string iniAi = _clsFun.FileIni("R", clsDefine.enuIni.Ini15AIAgent, "0");
                if (iniAi == "1" || iniAi == "S")
                {
                    ToolStripMenuItem mnuAI = new ToolStripMenuItem("Assistente AI");
                    mnuAI.Image = clsUiIcons.GetIcon("ai_agent", 16);
                    mnuAI.Font = new Font("Segoe UI", 9, FontStyle.Bold);
                    mnuAI.Alignment = ToolStripItemAlignment.Right;
                    mnuAI.Click += BtnAI_Click;
                    this.menuStrip1.Items.Add(mnuAI);
                }

                ApplyModernUi();
                LoadDashboard();
                AddButtonNumbers();

                lblGiorno.Text = DateTime.Now.ToString("dddd, dd MMMM yyyy", new System.Globalization.CultureInfo("it-IT"));
                lblClock.Text = DateTime.Now.ToString("HH:mm:ss");

                string iniDash = _clsFun.FileIni("R", clsDefine.enuIni.Ini16DashBoard, "1");
                if (iniDash == "0")
                {
                    pnlDash.Visible = false;
                    button16.Text = "«";
                    this.Width = 755;
                }
                else
                {
                    pnlDash.Visible = true;
                    button16.Text = "»";
                    this.Width = 1020;
                }

                button1.Select();
            }
            else
            {
                // Fallback di sicurezza se caricato senza splash
                Esci();
            }
        }

        private void ApplyModernUi()
        {
            try
            {
                // Top Menu Strip
                if (menuStrip1 != null)
                {
                    menuStrip1.Renderer = clsUiIcons.GetModernMenuRenderer();
                    if (esciToolStripMenuItem != null)
                        esciToolStripMenuItem.Image = clsUiIcons.GetIcon("exit", 16);
                }

                // 1: Anagrafica articoli (button1) - Blue Primary
                clsUiIcons.StyleMainMenuButton(button1, 1, "Anagrafica articoli", "articles", 38,
                    Color.FromArgb(238, 246, 255), Color.FromArgb(198, 220, 252), Color.FromArgb(96, 165, 250),
                    Color.FromArgb(30, 58, 138), Color.FromArgb(29, 78, 216));

                // 2: Anagrafica fornitori (button3) - Emerald Green
                clsUiIcons.StyleMainMenuButton(button3, 2, "Anagrafica fornitori", "suppliers", 38,
                    Color.FromArgb(236, 253, 245), Color.FromArgb(167, 243, 208), Color.FromArgb(52, 211, 153),
                    Color.FromArgb(6, 78, 59), Color.FromArgb(4, 120, 87));

                // 3: Gestione documenti (button6) - Amber Gold
                clsUiIcons.StyleMainMenuButton(button6, 3, "Gestione documenti", "documents_stack", 38,
                    Color.FromArgb(255, 253, 245), Color.FromArgb(254, 230, 138), Color.FromArgb(251, 191, 36),
                    Color.FromArgb(120, 53, 15), Color.FromArgb(217, 119, 6));

                // 4: Anagrafica clienti (button4) - Rose Crimson
                clsUiIcons.StyleMainMenuButton(button4, 4, "Anagrafica\nclienti", "clients_group", 32,
                    Color.FromArgb(255, 241, 242), Color.FromArgb(254, 205, 211), Color.FromArgb(251, 113, 133),
                    Color.FromArgb(136, 19, 55), Color.FromArgb(225, 29, 72));

                // 5: Anagrafica tessere (button10) - Violet Purple
                clsUiIcons.StyleMainMenuButton(button10, 5, "Anagrafica\ntessere", "fidelity_card", 32,
                    Color.FromArgb(245, 243, 255), Color.FromArgb(221, 214, 254), Color.FromArgb(167, 139, 250),
                    Color.FromArgb(76, 29, 149), Color.FromArgb(124, 58, 237));

                // 6: Divulgazioni da fornitori (button8) - Teal Cyan
                clsUiIcons.StyleMainMenuButton(button8, 6, "Divulgazioni\nda fornitori", "cloud_sync", 30,
                    Color.FromArgb(240, 253, 250), Color.FromArgb(153, 246, 228), Color.FromArgb(45, 212, 191),
                    Color.FromArgb(19, 78, 74), Color.FromArgb(13, 148, 136));

                // 7: Ordini a fornitori (button11) - Cyan Blue
                clsUiIcons.StyleMainMenuButton(button11, 7, "Ordini a\nfornitori", "orders_cart", 30,
                    Color.FromArgb(240, 249, 255), Color.FromArgb(186, 230, 253), Color.FromArgb(56, 189, 248),
                    Color.FromArgb(12, 74, 110), Color.FromArgb(2, 132, 199));

                // 8: Gestione offerte (button5) - Coral Orange
                clsUiIcons.StyleMainMenuButton(button5, 8, "Gestione\nofferte", "offers_promo", 30,
                    Color.FromArgb(255, 247, 237), Color.FromArgb(254, 215, 170), Color.FromArgb(251, 146, 60),
                    Color.FromArgb(124, 45, 18), Color.FromArgb(234, 88, 12));

                // 9: Inventario (button12) - Sky Blue
                clsUiIcons.StyleMainMenuButton(button12, 9, "Inventario", "inventory_warehouse", 30,
                    Color.FromArgb(239, 246, 255), Color.FromArgb(191, 219, 254), Color.FromArgb(96, 165, 250),
                    Color.FromArgb(30, 58, 138), Color.FromArgb(37, 99, 235));

                // 10: Variazioni POS/Bilance/Etichette (button7) - Indigo
                clsUiIcons.StyleMainMenuButton(button7, 10, "Variazioni\nPOS/Bilance/Eti", "pos_balance", 30,
                    Color.FromArgb(238, 242, 255), Color.FromArgb(199, 210, 254), Color.FromArgb(129, 140, 248),
                    Color.FromArgb(49, 46, 129), Color.FromArgb(79, 70, 229));

                // 11: Chiusura fine giornata (button9) - Slate Steel
                clsUiIcons.StyleMainMenuButton(button9, 11, "Chiusura fine\ngiornata", "day_closing", 30,
                    Color.FromArgb(248, 250, 252), Color.FromArgb(203, 213, 225), Color.FromArgb(148, 163, 184),
                    Color.FromArgb(15, 23, 42), Color.FromArgb(51, 65, 85));

                // 12: Utility di Import XLS (button13) - Excel Green
                clsUiIcons.StyleMainMenuButton(button13, 12, "Utility Import XLS", "excel_import", 22,
                    Color.FromArgb(240, 253, 244), Color.FromArgb(187, 247, 208), Color.FromArgb(74, 222, 128),
                    Color.FromArgb(20, 83, 45), Color.FromArgb(22, 163, 74));

                // 13: Attività (button14) - Warm Amber
                clsUiIcons.StyleMainMenuButton(button14, 13, "Attività", "tasks_prints", 22,
                    Color.FromArgb(254, 252, 232), Color.FromArgb(254, 240, 138), Color.FromArgb(250, 204, 21),
                    Color.FromArgb(113, 63, 18), Color.FromArgb(202, 138, 4));

                // 14: Statistiche (button15) - Cobalt Blue
                clsUiIcons.StyleMainMenuButton(button15, 14, "Statistiche", "statistics", 22,
                    Color.FromArgb(239, 246, 255), Color.FromArgb(191, 219, 254), Color.FromArgb(96, 165, 250),
                    Color.FromArgb(30, 58, 138), Color.FromArgb(29, 78, 216));

                // 15: Utilità (button2) - Titanium Slate
                clsUiIcons.StyleMainMenuButton(button2, 15, "Utilità", "tools_gear", 22,
                    Color.FromArgb(241, 245, 249), Color.FromArgb(203, 213, 225), Color.FromArgb(148, 163, 184),
                    Color.FromArgb(30, 41, 59), Color.FromArgb(71, 85, 105));

                // Form background & Dashboard styling
                this.BackColor = Color.FromArgb(243, 244, 246);
                if (pnlDash != null)
                {
                    pnlDash.BackColor = Color.FromArgb(248, 250, 252);
                    if (lblDashTitle != null)
                    {
                        lblDashTitle.BackColor = Color.FromArgb(30, 41, 59);
                        lblDashTitle.ForeColor = Color.White;
                    }
                    if (lblGiorno != null)
                    {
                        lblGiorno.BackColor = Color.FromArgb(241, 245, 249);
                        lblGiorno.ForeColor = Color.FromArgb(51, 65, 85);
                    }
                    if (lblClock != null)
                    {
                        lblClock.BackColor = Color.FromArgb(15, 23, 42);
                        lblClock.ForeColor = Color.FromArgb(16, 185, 129);
                    }
                    if (lblVarPos != null)
                    {
                        lblVarPos.Cursor = Cursors.Hand;
                        lblVarPos.Click += (s, ev) => button7.PerformClick();
                    }
                    if (lblVarEti != null)
                    {
                        lblVarEti.Cursor = Cursors.Hand;
                        lblVarEti.Click += (s, ev) => button7.PerformClick();
                    }
                    if (lblOfferte != null)
                    {
                        lblOfferte.Cursor = Cursors.Hand;
                        lblOfferte.Click += (s, ev) => button5.PerformClick();
                    }
                }
                if (button16 != null)
                {
                    button16.FlatStyle = FlatStyle.Flat;
                    button16.FlatAppearance.BorderSize = 0;
                    button16.BackColor = Color.FromArgb(51, 65, 85);
                    button16.ForeColor = Color.White;
                    button16.Cursor = Cursors.Hand;
                }
            }
            catch (Exception ex)
            {
                _clsFun.ErrorLog("ApplyModernUi", ex.Message);
            }
        }

        private void SetGraph()
        {
            DataRow[] j;
            string s = "SELECT * FROM TabFonts";
            DataTable t = _clsFun.FillTabSql("TabFonts", s, false, _strConSql);

            j = t.Select("tab_cod='001'");
            try
            {
                if (j.Length > 0 && ((string)j[0]["tab_col"]).Trim() != "")
                    this.BackColor = Color.FromName((string)j[0]["tab_col"]);
            }
            catch (Exception ex)
            {
                _clsFun.ErrorLog(ex.Message, Convert.ToString(j[0]["tab_cod"]));
            }

            foreach (Control c in this.Controls)
            {
                try
                {
                    if (c.GetType() == typeof(MenuStrip))
                    {
                        j = t.Select("tab_cod='002'");
                        if (j.Length > 0 && ((string)j[0]["tab_col"]).Trim() != "")
                        {
                            //c.BackColor = Color.FromName((string)j[0]["tab_col"]);
                            s = (string)j[0]["tab_col"];
                            c.BackColor = System.Drawing.ColorTranslator.FromHtml("#" + s);
                        }
                    }

                    if (c.GetType() == typeof(Button))
                    {
                        j = t.Select("tab_cod='003'");
                        if (j.Length > 0 && ((string)j[0]["tab_col"]).Trim() != "")
                            c.BackColor = Color.FromName((string)j[0]["tab_col"]);
                    }
                }
                catch (Exception ex)
                {
                    _clsFun.ErrorLog(ex.Message, c.Name);
                }
            }
        }

        private void frmAMain_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Escape)
            {
                if (Form.ActiveForm == this && Application.OpenForms.Count <= 1)
                {
                    if (MessageBox.Show("Confermi l'uscita da APOffice?", "Uscita da APOffice", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                    {
                        Esci();
                    }
                }
                e.Handled = true;
                e.SuppressKeyPress = true;
                return;
            }

            // Handle numeric shortcuts (1-9)
            if (e.KeyCode >= Keys.D1 && e.KeyCode <= Keys.D9)
            {
                ExecuteShortcut(e.KeyCode - Keys.D1);
            }
            // Handle numeric keypad (1-9)
            else if (e.KeyCode >= Keys.NumPad1 && e.KeyCode <= Keys.NumPad9)
            {
                ExecuteShortcut(e.KeyCode - Keys.NumPad1);
            }
            // Handle Function keys (F1-F12)
            else if (e.KeyCode >= Keys.F1 && e.KeyCode <= Keys.F12)
            {
                ExecuteShortcut(e.KeyCode - Keys.F1);
            }
        }

        private void ExecuteShortcut(int index)
        {
            if (_shortcutButtons != null && index >= 0 && index < _shortcutButtons.Count)
            {
                _shortcutButtons[index].PerformClick();
            }
        }
        private void esciToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Esci();
        }

        private void Esci()
        {
            this.Close();
        }

        private frmAIChat _chatForm = null;
        private void BtnAI_Click(object sender, EventArgs e)
        {
            if (_chatForm == null || _chatForm.IsDisposed)
            {
                _chatForm = new frmAIChat();
                // Posiziona il form a destra della finestra principale
                int x = this.Location.X + this.Width;
                int y = this.Location.Y;

                // Se esce dallo schermo, lo mettiamo a filo del bordo destro dello schermo
                Screen screen = Screen.FromControl(this);
                if (x + _chatForm.Width > screen.WorkingArea.Right)
                {
                    x = screen.WorkingArea.Right - _chatForm.Width;
                }

                _chatForm.Location = new Point(x, y);
            }

            if (!_chatForm.Visible)
                _chatForm.Show();
            else
                _chatForm.Activate();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            ShowForm(new frmAnaArticolo());
        }

        private void button2_Click(object sender, EventArgs e)
        {
            ShowForm(new frmMnuUtilita());
        }

        private void button3_Click(object sender, EventArgs e)
        {
            ShowForm(new frmSeekFor());
        }

        private void button4_Click(object sender, EventArgs e)
        {
            string s = _clsFun.ParGet(clsDefine.enuParametri.Par038Password, _strConSql);
            string[] a = s.Split('|');
            s = a[0].Trim();

            frmSeekCli fSeek = new frmSeekCli();
            if (s != "")
            {
                frmUser f = new frmUser();
                f._strConSql = _clsFun.ConSql("");
                f._strTip = "CLI";
                f._strPwd = s;
                if (f.ShowDialog() == DialogResult.OK || f._bolSta)
                    ShowForm(fSeek);
            }
            else
                ShowForm(fSeek);
        }

        private void button5_Click(object sender, EventArgs e)
        {
            ShowForm(new frmSeekOff());
        }

        private void button6_Click(object sender, EventArgs e)
        {
            ShowForm(new frmSeekDocs());
        }

        private void button7_Click(object sender, EventArgs e)
        {
            Boolean b = true;
            string sFil = "";
            string sTer = "";
            string s = _clsQry.SmfCtrl("ENTRA");

            string[] a = s.Split('|');

            if (a.Length > 1)
            {
                sFil = a[0];
                sTer = a[1];

                if (sTer != "")
                {
                    MessageBox.Show("Terminale " + sTer + " attivo in GESIONE VARIAZIONI, accesso negato!", "CONTROLLO UTENTI", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    b = false;
                }
            }

            if (b)
            {
                frmGesVariazioni f = new frmGesVariazioni();
                if (sTer != "")
                    f._strSta = "NOVAR";
                ShowForm(f);
            }

            if (File.Exists(sFil))
                File.Delete(sFil);
        }

        private void button8_Click(object sender, EventArgs e)
        {
            ShowForm(new frmGesForDivulgazioni());
        }

        private void button9_Click(object sender, EventArgs e)
        {
            const string sBat = @"c:\approject\CHIUSURA.BAT";
            if (!File.Exists(sBat))
            {
                MessageBox.Show("File CHIUSURA.BAT non trovato in c:\\approject\\", "CHIUSURA", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            try
            {
                Process p = new Process();
                p.StartInfo.FileName = sBat;
                p.StartInfo.WorkingDirectory = System.IO.Path.GetDirectoryName(sBat);
                p.StartInfo.WindowStyle = ProcessWindowStyle.Normal;
                p.StartInfo.UseShellExecute = true;
                p.Start();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Errore avvio CHIUSURA.BAT:\n" + ex.Message, "CHIUSURA", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void button10_Click(object sender, EventArgs e)
        {
            string s = _clsFun.ParGet(clsDefine.enuParametri.Par027Fidelity, _strConSql);
            if (s.Length > 0 && s.Substring(0, 1) == "2")
                ShowForm(new frmSeekTessere2());
            else
                ShowForm(new frmSeekTessere());
        }

        private void button11_Click(object sender, EventArgs e)
        {
            ShowForm(new frmSeekOrdFor());
        }

        private void button12_Click(object sender, EventArgs e)
        {
            ShowForm(new frmSeekInventario());
        }

        private void CtrlDivulgazioni()
        {
            string s = "";
            string sPath = Path.GetDirectoryName(_clsDef.FILEINI);

            if (_clsFun.ParGet(clsDefine.enuParametri.ParDivForDownLoad, _strConSql) == "S")
            {
                s = Path.GetDirectoryName(_clsDef.FILEINI) + "\\Div" + DateTime.Today.ToString("yyyyMMdd") + ".sem";

                if (!File.Exists(s))
                {
                    string[] sFils = Directory.GetFiles(sPath, "Div*.flg");
                    foreach (string sFil in sFils)
                    {
                        Console.WriteLine("aaaa");
                        if (!sFil.Equals(s))
                            File.Delete(sFil);
                    }
                    new frmGesForDivulgazioni().ForDownLoad();
                    File.Copy(_clsDef.FILEINI, s);
                }
            }

            s = Path.GetDirectoryName(_clsDef.FILEINI) + "\\Bck" + DateTime.Today.ToString("yyyyMMdd") + ".sem";

            if (!File.Exists(s))
            {
                //Backup
                if (File.Exists(_clsDef.BATCHBACKUP))
                {
                    Process p = new Process();
                    p.StartInfo.FileName = _clsDef.BATCHBACKUP;
                    p.StartInfo.Arguments = "-r";
                    p.StartInfo.ErrorDialog = true;
                    p.StartInfo.WindowStyle = ProcessWindowStyle.Minimized;
                    p.Start();
                    //p.WaitForExit(1000 * 60 * 2);    // wait up to 2 minutes.

                    s = Path.GetDirectoryName(_clsDef.FILEINI);
                    string[] sFils = Directory.GetFiles(s, "Bck*.sem");
                    foreach (string ss in sFils)
                        File.Delete(ss);

                    s = Path.GetDirectoryName(_clsDef.FILEINI) + "\\Bck" + DateTime.Today.ToString("yyyyMMdd") + ".sem";

                    File.Copy(_clsDef.FILEINI, s);
                }
            }

        }

        private void button13_Click(object sender, EventArgs e)
        {
            ShowForm(new frmUtyImportXls());
        }

        private void button14_Click(object sender, EventArgs e)
        {
            ShowForm(new frmMnuStampe());
        }

        private void button15_Click(object sender, EventArgs e)
        {
            ShowForm(new frmGesStat2());
        }

        private void button16_Click(object sender, EventArgs e)
        {
            if (pnlDash.Visible)
            {
                pnlDash.Visible = false;
                button16.Text = "«";
                this.Width = 755;
            }
            else
            {
                pnlDash.Visible = true;
                button16.Text = "»";
                this.Width = 1020;
            }
        }

        //private void button16_Click(object sender, EventArgs e)
        //{
        //    new clsMail().MailProva();
        //}

        //private void button16_Click(object sender, EventArgs e)
        //{
        //    char c = Convert.ToChar(65);
        //    string s = c.ToString();
        //    Console.WriteLine("aaaaaaaaa");
        //}

        private void AddButtonNumbers()
        {
            _shortcutButtons = new List<Button>()
            {
                button1,  // 1: Anagrafica articoli
                button3,  // 2: Anagrafica fornitori
                button6,  // 3: Gestione documenti
                button4,  // 4: Anagrafica clienti
                button10, // 5: Anagrafica tessere
                button8,  // 6: Divulgazioni da fornitori
                button11, // 7: Ordini a fornitori
                button5,  // 8: Gestione offerte
                button12, // 9: Inventario
                button7,  // 10: Variazioni POS/Bilance/Etichette
                button9,  // 11: Chiusura fine giornata
                button13, // 12: Utility di Import XLS
                button14, // 13: Attività
                button15, // 14: Statistiche
                button2   // 15: Utilità
            };

            // Rimuovi eventuali label residue
            foreach (Button btn in _shortcutButtons)
            {
                if (btn != null)
                {
                    List<Control> toRemove = new List<Control>();
                    foreach (Control sub in btn.Controls)
                    {
                        if (sub.Tag != null && sub.Tag.ToString() == "BtnNum")
                            toRemove.Add(sub);
                    }
                    foreach (Control sub in toRemove)
                        btn.Controls.Remove(sub);
                }
            }
        }

        private void Controlli()
        {

            DataTable t = _clsFun.FillTabSql("TabNegozi", "SELECT * FROM TabNegozi WHERE tab_ann=0 AND tab_tip='L'", false, _strConSql);
            if (t.Rows.Count == 0)
                MessageBox.Show("Tabella negozi non Configurata!", "CHIAMARE L'ASSISTENZA", MessageBoxButtons.OK, MessageBoxIcon.Error);
            else if (((string)t.Rows[0]["tab_pos"]).Trim() == "")
                MessageBox.Show("Casse non configurate in tabella negozi!", "CHIAMARE L'ASSISTENZA", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }

        /// <summary>
        /// Abilita button9 (Chiusura fine giornata) solo se esiste c:\approject\CHIUSURA.BAT.
        /// </summary>
        private void CtrlChiusuraBat()
        {
            const string sBat = @"c:\approject\CHIUSURA.BAT";
            button9.Enabled = File.Exists(sBat);
        }

        /// <summary>
        /// Controllo automatico offerte all'avvio: attiva offerte correnti, chiude le scadute
        /// e genera le variazioni POS/Bilance ed Etichette.
        /// </summary>
        private void CtrlOfferte()
        {
            try
            {
                int p = 0;
                int e = 0;
                new clsVariazioni().ConsolidaOfferteAllAvvio(out p, out e);
            }
            catch (Exception ex)
            {
                _clsFun.ErrorLog("frmAMain.CtrlOfferte", ex.Message);
            }
        }

        private void ShowForm(Form f)
        {
            try
            {
                frmWait.ShowWait(f.Text);

                // Chiude frmWait appena il nuovo form viene mostrato a video
                f.Shown += (s, ev) =>
                {
                    f.BeginInvoke(new Action(() =>
                    {
                        frmWait.CloseWait();

                        Control curActive = f.ActiveControl;
                        f.BringToFront();
                        f.Activate();
                        SetForegroundWindow(f.Handle);
                        if (curActive != null && curActive.CanFocus)
                        {
                            curActive.Focus();
                        }
                        else
                        {
                            f.Focus();
                        }

                        // Assicuriamoci che la finestra principale non rubi il focus
                        this.SendToBack();
                    }));
                };

                string iniModal = _clsFun.FileIni("R", clsDefine.enuIni.Ini17Modal, "1");
                if (iniModal == "0")
                {
                    f.Show(this);
                }
                else
                {
                    f.ShowDialog(this);
                    f.Dispose();
                }
            }
            catch (Exception ex)
            {
                _clsFun.ErrorLog(ex.Message, "ShowForm: " + f.Name);
                frmWait.CloseWait();
            }
            finally
            {
                // Fallback nel caso in cui Shown non venga chiamato o per ShowDialog conclusi
                frmWait.CloseWait();
                LoadDashboard();
            }
        }
        private void LoadDashboard()
        {
            try
            {
                // Offerte (Moved to top to ensure it runs even if other queries fail)
                string sSqlOff = @"/* DASH_OFF_V5 */ 
                                  SELECT 
                                    COUNT(*) as total,
                                    SUM(CASE WHEN oft_sta='A' OR oft_sta='D' OR (oft_sta='N' AND CONVERT(date, oft_dti) <= CONVERT(date, GETDATE()) AND CONVERT(date, oft_dtf) >= CONVERT(date, GETDATE())) THEN 1 ELSE 0 END) as active
                                  FROM GesOffTestate";
                DataTable tOff = _clsFun.FillTabSql("GesOffTestate", sSqlOff, false, _strConSql);
                if (tOff != null && tOff.Rows.Count > 0)
                {
                    int tot = (tOff.Rows[0]["total"] != DBNull.Value) ? Convert.ToInt32(tOff.Rows[0]["total"]) : 0;
                    int act = (tOff.Rows[0]["active"] != DBNull.Value) ? Convert.ToInt32(tOff.Rows[0]["active"]) : 0;
                    lblOfferte.Text = "Offerte\n" + act + " su " + tot;
                    lblOfferte.ForeColor = (act > 0) ? Color.FromArgb(192, 57, 43) : Color.FromArgb(44, 62, 80);
                }

                // Articoli
                DataTable t = _clsFun.FillTabSql("AnaArticoli", "SELECT COUNT(*) as tot FROM AnaArticoli", false, _strConSql);
                if (t.Rows.Count > 0) lblArticoli.Text = "Articoli\n" + t.Rows[0]["tot"].ToString();

                // Barcode
                t = _clsFun.FillTabSql("AnaBarcode", "SELECT COUNT(*) as tot FROM AnaBarcode", false, _strConSql);
                if (t.Rows.Count > 0) lblBarcode.Text = "Barcode\n" + t.Rows[0]["tot"].ToString();

                // Fornitori
                t = _clsFun.FillTabSql("AnaFornitori", "SELECT COUNT(*) as tot FROM AnaFornitori WHERE for_ann=0", false, _strConSql);
                if (t.Rows.Count > 0) lblFornitori.Text = "Fornitori\n" + t.Rows[0]["tot"].ToString();

                // Clienti
                t = _clsFun.FillTabSql("AnaClienti", "SELECT COUNT(*) as tot FROM AnaClienti WHERE cli_ann=0", false, _strConSql);
                if (t.Rows.Count > 0) lblClienti.Text = "Clienti\n" + t.Rows[0]["tot"].ToString();

                // Fidelity
                t = _clsFun.FillTabSql("AnaTessere", "SELECT COUNT(*) as tot FROM AnaTessere", false, _strConSql);
                if (t.Rows.Count > 0) lblFidelity.Text = "Fidelity\n" + t.Rows[0]["tot"].ToString();

                // Ultimo aggiornamento
                t = _clsFun.FillTabSql("GesVariazioni", "SELECT MAX(var_dti) as ult FROM GesVariazioni", false, _strConSql);
                if (t != null && t.Rows.Count > 0 && !DBNull.Value.Equals(t.Rows[0]["ult"]))
                    lblUltimoAgg.Text = "Ultimo agg. listini:\n" + Convert.ToDateTime(t.Rows[0]["ult"]).ToString("dd/MM/yyyy HH:mm");

                // Variazioni POS (Ready to send)
                t = _clsFun.FillTabSql("GesVariazioni", "SELECT COUNT(DISTINCT v.var_art) as tot FROM GesVariazioni v INNER JOIN AnaArticoli a ON LTRIM(RTRIM(v.var_art)) = a.art_cod WHERE v.var_tip='POS' AND v.var_num='000'", false, _strConSql);
                if (t != null && t.Rows.Count > 0)
                {
                    int n = Convert.ToInt32(t.Rows[0]["tot"]);
                    lblVarPos.Text = "Var. POS\n" + n.ToString("N0");
                    lblVarPos.ForeColor = (n > 0) ? Color.FromArgb(230, 126, 34) : Color.FromArgb(44, 62, 80);
                }

                // Etichette (To print)
                t = _clsFun.FillTabSql("GesVariazioni", "SELECT COUNT(DISTINCT v.var_art) as tot FROM GesVariazioni v INNER JOIN AnaArticoli a ON LTRIM(RTRIM(v.var_art)) = a.art_cod WHERE v.var_tip='ETI' AND v.var_num='000'", false, _strConSql);
                if (t != null && t.Rows.Count > 0)
                {
                    int n = Convert.ToInt32(t.Rows[0]["tot"]);
                    lblVarEti.Text = "Etichette\n" + n.ToString("N0");
                    lblVarEti.ForeColor = (n > 0) ? Color.FromArgb(230, 126, 34) : Color.FromArgb(44, 62, 80);
                }
            }
            catch (Exception ex)
            {
                _clsFun.ErrorLog(ex.Message, "LoadDashboard");
            }
        }

        private void tmrDash_Tick(object sender, EventArgs e)
        {
            lblClock.Text = DateTime.Now.ToString("HH:mm:ss");

            // Aggiornamento periodico dei dati (ogni 30 secondi)
            if (DateTime.Now.Second % 30 == 0)
            {
                LoadDashboard();
                lblGiorno.Text = DateTime.Now.ToString("dddd, dd MMMM yyyy", new System.Globalization.CultureInfo("it-IT"));
            }
        }
    }
}
