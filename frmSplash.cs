using System;
using System.Data;
using System.Drawing;
using System.IO;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace APOffice
{
    public partial class frmSplash : Form
    {
        clsFuncs _clsFun = new clsFuncs();
        clsDefine _clsDef = new clsDefine();
        string _strConSql = "";

        public frmSplash()
        {
            InitializeComponent();
        }

        private async void frmSplash_Load(object sender, EventArgs e)
        {
            try
            {
                // Carica il logo specifico richiesto dall'utente
                string logoPath = @"C:\ApProject\Temp\Img\Loghi\artLogo.png";
                if (!File.Exists(logoPath))
                    logoPath = Path.Combine(Application.StartupPath, "Report\\LogoSplash.png");

                if (!File.Exists(logoPath))
                    logoPath = Path.Combine(Application.StartupPath, "Report\\LogoFattura.png");

                if (File.Exists(logoPath))
                {
                    using (FileStream fs = new FileStream(logoPath, FileMode.Open, FileAccess.Read))
                    {
                        imgLogo.Image = Image.FromStream(fs);
                    }
                }
                string appVer = System.Reflection.Assembly.GetExecutingAssembly().GetName().Version.ToString();
                lblTitle.Text = "APOFFICE SUITE v" + appVer;
            }
            catch { }

            // Attendi che l'interfaccia sia pronta
            await Task.Delay(700);

            // Avvia l'inizializzazione in background
            bool success = await Task.Run(() => InitProcess());

            if (success)
            {
                // Un ulteriore secondo per "WOW effect" e stabilità
                await Task.Delay(1000);
                this.DialogResult = DialogResult.OK;
                this.Close();
            }
            else
            {
                // Errore: ferma l'animazione e mostra il pulsante ESCI
                UpdateStatus("ERRORE CRITICO ALL'AVVIO");
                pbLoading.Style = ProgressBarStyle.Blocks;
                pbLoading.Value = 100;
                pbLoading.ForeColor = Color.Red;
                btnExit.Visible = true;
                lblStatus.ForeColor = Color.FromArgb(231, 76, 60);
            }
        }

        private bool InitProcess()
        {
            try
            {
                UpdateStatus("Verifica parametri di configurazione...");
                System.Threading.Thread.Sleep(600);

                _strConSql = _clsFun.ConSql("");
                if (string.IsNullOrEmpty(_strConSql))
                {
                    UpdateStatus("ERRORE: Stringa di connessione SQL mancante in APOffice.ini");
                    return false;
                }

                UpdateStatus("Connessione al database SQL Server...");
                // Test connessione
                using (System.Data.SqlClient.SqlConnection cn = new System.Data.SqlClient.SqlConnection(_strConSql))
                {
                    cn.Open();
                    UpdateStatus("Database connesso. Verifica schema...");
                    cn.Close();
                }
                System.Threading.Thread.Sleep(400);

                UpdateStatus("Esecuzione aggiornamenti database...");
                new clsDbMigration().CheckDb(_strConSql);
                System.Threading.Thread.Sleep(400);

                UpdateStatus("Verifica licenza e azienda...");
                string s = _clsFun.FileIni("R", clsDefine.enuIni.Ini09CodiceAzienda, "");
                if (s == "") s = "001";

                DataTable t = new clsQuery().ConfAzienda(s);
                if (t.Rows.Count == 0)
                {
                    UpdateStatus("ERRORE: Azienda '" + s + "' non trovata o configurazione corrotta.");
                    return false;
                }

                UpdateStatus("Caricamento moduli interfaccia...");
                System.Threading.Thread.Sleep(400);

                UpdateStatus("Pronto.");
                return true;
            }
            catch (Exception ex)
            {
                UpdateStatus("ERRORE CRITICO: " + ex.Message);
                return false;
            }
        }

        private void UpdateStatus(string text)
        {
            if (this.InvokeRequired)
            {
                this.Invoke(new Action<string>(UpdateStatus), text);
                return;
            }
            lblStatus.Text = text;
        }

        private void btnExit_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.Abort;
            this.Close();
        }
    }
}
