using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using System.Diagnostics;
using System.Threading;

namespace APOffice
{
    public partial class frmGesForDivulgazioni : Form
    {
        clsDefine _clsDef = new clsDefine();
        clsFuncs _clsFun = new clsFuncs();
        clsQuery _clsQry = new clsQuery();
        clsVariazioni _clsVar = new clsVariazioni();

        private const string TABANAART = "AnaArticoli";
        private const string TABANAEAN = "AnaBarcode";
        private const string TABLISACQ = "GesLisAcquisto";
        private const string TABLISVEN = "GesLisVendita";
        private const string TABGESOFT = "GesOffTestate";
        private const string TABGESOFA = "GesOffArticoli";
        private const string TABTABFOR = "TabTabFornitori";
        private const string TABTABIVA = "TabIva";
        private const string TABANAOFF = "GesOffArticoli";
        private const string TABGESFAT = "GesFatTestate";
        private const string TABGESMOT = "GesMovTestate";
        private const string TABGESMOV = "GesMovimenti";
        private const string TABTMPDIV = "TmpDiv";

        //private const string TABTMPFDR = "ForDivRig";
        //private const string TABTMPEAN = "ForDivEan";
        private const string TABTMPFDR = "GesDivForArticoli";
        private const string TABTMPEAN = "GesDivForBarcode";

        private const string VARART = "ART";
        private const string VARFAT = "FAT";
        private const string VARNCA = "NAC";
        private const string VAROFF = "OFF";

        private string _strConSql = "";
        private string _strConSqlMdb = "";
        private string _strConSqlLog = "";
        private string _str014AggPVenditaSoloFattura = "";
        private string _str015AggDesArtFromFornitore = "";
        private string _strPar017EcrDivFornitori = "";
        private string _strPar020EsclusioneEan2021davDiv = "";

        private string _strForRif = "";
        private string _strRepDef = "";

        private DataSet _dasGen = new DataSet();

        public frmGesForDivulgazioni()
        {
            InitializeComponent();
            ApplyMenuIcons();
            new clsGesGraph().SetGraph(this, 0);
            _strConSql = _clsFun.ConSql("");
            _strConSqlLog = _clsFun.ConSql("4");
            _strConSqlMdb = _clsFun.ConMdb(Path.GetDirectoryName(_clsDef.FILEINI) + "\\ApForTracciati.mdb");
        }

        private void ApplyMenuIcons()
        {
            // Set Font for better emoji support
            menuStrip1.Font = new Font("Segoe UI", 9F);

            esciToolStripMenuItem.Text = "🚪 Esci";
            downloadToolStripMenuItem.Text = "📥 Download";
            importaToolStripMenuItem.Text = "📂 Carica";
            elaboraToolStripMenuItem.Text = "🔄 Aggiorna";
            stampaEtichetteToolStripMenuItem.Text = "🏷️ Generazione etichette";
            stampeToolStripMenuItem.Text = "🖨️ Stampe";
            
            articoliNuoviToolStripMenuItem.Text = "🆕 Articoli nuovi/annullati";
            differenzeToolStripMenuItem.Text = "⚖️ Differenze anagrafiche";
            stampaDocumentoToolStripMenuItem.Text = "📄 Stampa documento";
            
            utilityToolStripMenuItem.Text = "🛠️ Utility";
            conversioneTabelleToolStripMenuItem.Text = "📊 Conversione tabelle";
            attivaArticoliDelDocumentoToolStripMenuItem.Text = "✅ Attiva articoli del documento";
        }

        private void frmGesForDivulgazioni_Load(object sender, EventArgs e)
        {
            this.Show();

            _strForRif = _clsFun.ParGet(clsDefine.enuParametri.ParForDefault, _strConSql);
            if (_strForRif.Length > 5)
                _strForRif = _strForRif.Substring(0, 5);

            _strRepDef = _clsFun.ParGet(clsDefine.enuParametri.Par013CodRepxDefault, _strConSql);

            _str014AggPVenditaSoloFattura = _clsFun.ParGet(clsDefine.enuParametri.Par014AggPVenditaSoloFattura, _strConSql);
            _str015AggDesArtFromFornitore = _clsFun.ParGet(clsDefine.enuParametri.Par015AggDesArtFromFornitore, _strConSql);
            _strPar017EcrDivFornitori = _clsFun.ParGet(clsDefine.enuParametri.Par017EcrDivFornitori, _strConSql);
            _strPar020EsclusioneEan2021davDiv = _clsFun.ParGet(clsDefine.enuParametri.Par020EsclusioneEan2021davDiv, _strConSql);

            menuStrip1.Enabled = false;

            SetDgv1();
            SetDgv2();
            FillTab();
            ForImport();

            menuStrip1.Enabled = true;
        }

        private void frmGesForDivulgazioni_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Escape)
                Esci();
        }
        private void esciToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Esci();
        }
        private void Esci()
        {
            if (dgv1 != null && ((DataTable)dgv1.DataSource).Rows.Count > 0)
            {
                if (MessageBox.Show("Uscendo vengono perse eventuali modifiche, confermi?", this.Text, MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                {
                    this.Close();
                    this.Dispose();
                }
            }
            else
                this.Close();
        }

        private void importaToolStripMenuItem_Click(object sender, EventArgs e)
        {
            ForImport();
            //MessageBox.Show("Fine caricamento!");
        }

        private void stampaEtichetteToolStripMenuItem_Click(object sender, EventArgs e)
        {
            CurrencyManager cm = dgv1.BindingContext[dgv1.DataSource, dgv1.DataMember] as CurrencyManager;
            if (cm.Position >= 0)
            {
                DataRowView r = cm.List[dgv1.CurrentRow.Index] as DataRowView;
                DataRow x = r.Row;

                string sTva = (string)x["dit_tva"];
                string sDdo = "";

                if (!DBNull.Value.Equals(x["dit_ddo"]))
                    sDdo = ((DateTime)x["dit_ddo"]).ToString("dd/MM/yyy");
                string sNum = (string)x["dit_num"];

                string s = "Confermi la generazione etichette per ";
                if (sTva == "FAT")
                {
                    s += "la fattura ";
                    s += " del " + sDdo + "?";
                }
                else
                    s += "le variazioni ";

                if (MessageBox.Show(s, this.Text, MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                {
                    DataTable t = _dasGen.Tables[TABTMPFDR];
                    DataView v = new DataView(t, "div_tva='" + sTva + "' AND div_num='" + sNum + "'", "", DataViewRowState.CurrentRows);
                    foreach (DataRowView rr in v)
                        rr["div_eti"] = "S";
                }
            }
        }

        private void SetDgv1()
        {
            DataTable tNeg = _clsFun.FillTabSql("TabNegozi", "SELECT * FROM TabNegozi WHERE tab_ann=0 ORDER BY tab_cod", false, _strConSql);

            dgv1.AutoGenerateColumns = false;
            //dgv1.VirtualMode = true;
            //dgv1.Dock = DockStyle.Fill;
            dgv1.AllowUserToAddRows = false;
            dgv1.ReadOnly = false;
            dgv1.AllowUserToDeleteRows = false;

            DataGridViewTextBoxColumn cTbc;
            //DataGridViewCheckBoxColumn cCbc;
            DataGridViewComboBoxColumn cCmb;

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "dit_cho";
            cTbc.Name = "S";
            cTbc.Width = 25;
            cTbc.ValueType = typeof(string);
            cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            cTbc.ReadOnly = true;
            cTbc.ToolTipText = "Elabora, Parcheggia, Cancella";
            dgv1.Columns.Add(cTbc);

            //cTbc = new DataGridViewTextBoxColumn();
            //cTbc.DataPropertyName = "dit_neg";
            //cTbc.Name = "Negozio";
            //cTbc.Width = 35;
            //cTbc.ValueType = typeof(string);
            ////cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            //cTbc.ReadOnly = true;
            //dgv1.Columns.Add(cTbc);

            cCmb = new DataGridViewComboBoxColumn();
            cCmb.DataPropertyName = "dit_neg";
            cCmb.Name = "Negozio";
            cCmb.Width = 50;
            cCmb.DataSource = tNeg;
            cCmb.ValueMember = "tab_cod";
            cCmb.DisplayMember = "tab_cod";
            cCmb.DisplayStyle = DataGridViewComboBoxDisplayStyle.ComboBox;
            dgv1.Columns.Add((DataGridViewColumn)cCmb);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "dit_num";
            cTbc.Name = "Numero";
            cTbc.Width = 35;
            cTbc.ValueType = typeof(string);
            //cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            cTbc.ReadOnly = true;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "dit_fil";
            cTbc.Name = "File";
            cTbc.Width = 100;
            cTbc.ValueType = typeof(string);
            //cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            cTbc.ReadOnly = true;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "dit_tva";
            cTbc.Name = "Tipo variazione";
            cTbc.Width = 50;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = true;
            //cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "dit_des";
            cTbc.Name = "Descrizione";
            cTbc.Width = 100;
            cTbc.ValueType = typeof(string);
            //cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            cTbc.ReadOnly = true;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "dit_soi";
            cTbc.Name = "Dal";
            cTbc.Width = 80;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = true;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "dit_sof";
            cTbc.Name = "al";
            cTbc.Width = 80;
            cTbc.ValueType = typeof(string);
            cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            cTbc.ReadOnly = true;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "dit_ndo";
            cTbc.Name = "Numero";
            cTbc.Width = 80;
            cTbc.ValueType = typeof(string);
            cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            cTbc.ReadOnly = true;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "dit_ddo";
            cTbc.Name = "Data";
            cTbc.Width = 70;
            cTbc.ValueType = typeof(string);
            cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            cTbc.ReadOnly = false;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "dit_pth";
            cTbc.Name = "PATH";
            cTbc.Width = 300;
            cTbc.ValueType = typeof(string);
            //cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            cTbc.ReadOnly = false;
            dgv1.Columns.Add(cTbc);
        }

        private void SetDgv2()
        {
            dgv2.AutoGenerateColumns = false;
            //dgv2.VirtualMode = true;
            //dgv2.Dock = DockStyle.Fill;
            dgv2.AllowUserToAddRows = false;
            dgv2.ReadOnly = false;
            dgv2.AllowUserToDeleteRows = false;

            DataGridViewTextBoxColumn cTbc;
            //DataGridViewCheckBoxColumn cCbc;
            //DataGridViewComboBoxColumn cCmb;

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "div_nri";
            cTbc.Name = "Riga";
            cTbc.Width = 38;
            cTbc.ValueType = typeof(string);
            //cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            cTbc.ReadOnly = true;
            dgv2.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "div_var";
            cTbc.Name = "Variazione";
            cTbc.Width = 30;
            cTbc.ValueType = typeof(string);
            //cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            cTbc.ReadOnly = true;
            cTbc.ToolTipText = "No aggiorna";
            dgv2.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "div_art";
            cTbc.Name = "Articolo";
            cTbc.Width = 50;
            cTbc.ValueType = typeof(string);
            //cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            cTbc.ReadOnly = true;
            dgv2.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "div_arf";
            cTbc.Name = "Art.fornitore";
            cTbc.Width = 40;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = true;
            //cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dgv2.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "div_ard";
            cTbc.Name = "Descrizione";
            cTbc.Width = 180;
            cTbc.ValueType = typeof(string);
            //cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            cTbc.ReadOnly = true;
            cTbc.ToolTipText = "Doppio click per anagrafica";
            dgv2.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "div_iva";
            cTbc.Name = "IVA";
            cTbc.Width = 30;
            cTbc.ValueType = typeof(string);
            //cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            cTbc.ReadOnly = true;
            dgv2.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "div_qta";
            cTbc.Name = "Q.tà";
            cTbc.Width = 40;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = true;
            cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            cTbc.DefaultCellStyle.Format = "###,##0.00";
            dgv2.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "art_cos";
            cTbc.Name = "Costo anagrafico";
            cTbc.Width = 40;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = true;
            cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            cTbc.DefaultCellStyle.Format = "###,##0.000";
            dgv2.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "div_cos";
            cTbc.Name = "Costo";
            cTbc.Width = 40;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = true;
            cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            cTbc.DefaultCellStyle.Format = "###,##0.000";
            dgv2.Columns.Add(cTbc);
            dgv2.Columns[cTbc.Name].DefaultCellStyle.BackColor = Color.Gold;

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "dif_cos";
            cTbc.Name = "Differenza costi";
            cTbc.Width = 40;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = true;
            cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            cTbc.DefaultCellStyle.Format = "###,##0.000";
            dgv2.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "div_imp";
            cTbc.Name = "Valore";
            cTbc.Width = 40;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = true;
            cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            cTbc.DefaultCellStyle.Format = "###,##0.000";
            dgv2.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "art_prp";
            cTbc.Name = "Prezzo anagrafico";
            cTbc.Width = 40;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = true;
            cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            cTbc.DefaultCellStyle.Format = "###,##0.00";
            dgv2.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "art_mav";
            cTbc.Name = "Margine";
            cTbc.Width = 40;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = true;
            cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            cTbc.DefaultCellStyle.Format = "###,##0.00";
            dgv2.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "div_prp";
            cTbc.Name = "Prezzo consigliato";
            cTbc.Width = 40;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = true;
            cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            cTbc.DefaultCellStyle.Format = "###,##0.00";
            dgv2.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "div_mav";
            cTbc.Name = "Margine";
            cTbc.Width = 40;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = true;
            cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            cTbc.DefaultCellStyle.Format = "###,##0.00";
            dgv2.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "var_prp";
            cTbc.Name = "Nuovo Prezzo";
            cTbc.Width = 40;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = false;
            cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            cTbc.DefaultCellStyle.Format = "###,##0.00";
            dgv2.Columns.Add(cTbc);
            dgv2.Columns[cTbc.Name].DefaultCellStyle.BackColor = Color.GreenYellow;

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "var_mav";
            cTbc.Name = "Margine";
            cTbc.Width = 40;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = false;
            cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            cTbc.DefaultCellStyle.Format = "###,##0.00";
            dgv2.Columns.Add(cTbc);

            //cTbc = new DataGridViewTextBoxColumn();
            //cTbc.DataPropertyName = "for_tgr";
            //cTbc.Name = "T.gramm";
            //cTbc.Width = 30;
            //cTbc.ValueType = typeof(string);
            ////cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            //cTbc.ReadOnly = true;
            //dgv2.Columns.Add(cTbc);

            //cTbc = new DataGridViewTextBoxColumn();
            //cTbc.DataPropertyName = "for_umi";
            //cTbc.Name = "Peso netto";
            //cTbc.Width = 40;
            //cTbc.ValueType = typeof(string);
            //cTbc.ReadOnly = false;
            //cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            //cTbc.DefaultCellStyle.Format = "###,##0.00";
            //dgv2.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "div_sta";
            cTbc.Name = "Stato";
            cTbc.Width = 20;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = false;
            dgv2.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "div_eti";
            cTbc.Name = "Etichetta";
            cTbc.Width = 20;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = false;
            cTbc.ToolTipText = "Crea etichetta";
            dgv2.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "liv_lis";
            cTbc.Name = "Listino di vendita";
            cTbc.Width = 50;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = false;
            cTbc.ToolTipText = "001 negozio";
            dgv2.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "div_stf";
            cTbc.Name = "Stato fornitore";
            cTbc.Width = 20;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = false;
            dgv2.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "div_off";
            cTbc.Name = "Offerta";
            cTbc.Width = 150;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = true;
            dgv2.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "div_rep";
            cTbc.Name = "Reparto";
            cTbc.Width = 50;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = true;
            dgv2.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "div_umi";
            cTbc.Name = "UM";
            cTbc.Width = 30;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = true;
            dgv2.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "div_tgr";
            cTbc.Name = "t.grammatura";
            cTbc.Width = 30;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = true;
            dgv2.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "div_pne";
            cTbc.Name = "Peso netto";
            cTbc.Width = 60;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = true;
            cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dgv2.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "div_pxc";
            cTbc.Name = "Pxc";
            cTbc.Width = 30;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = true;
            cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dgv2.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "div_af2";
            cTbc.Name = "Articolo fornitore consegnatario";
            cTbc.Width = 60;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = false;
            dgv2.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "for_ecr";
            cTbc.Name = "ECR fornitore";
            cTbc.Width = 50;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = false;
            dgv2.Columns.Add(cTbc);
        }

        private void newFillTab()
        {
            dgv1.DataSource = new clsGenTabTmp().TabTmpForDiv("ForDivTes");
            DataTable t = new clsGenTabTmp().TabTmpForDiv(TABTMPFDR);
            DataColumn[] keys = new DataColumn[4];
            keys[0] = t.Columns["div_tva"];
            keys[1] = t.Columns["div_num"];
            keys[2] = t.Columns["div_arf"];
            keys[3] = t.Columns["div_nri"];
            t.PrimaryKey = keys;
            _dasGen.Tables.Add(t);

            t = new clsGenTabTmp().TabTmpForDiv(TABTMPEAN);
            _dasGen.Tables.Add(t);

            string s = "SELECT * FROM TabIva";
            DataTable tIva = _clsFun.FillTabSql(TABTABIVA, s, false, _strConSql);
            _dasGen.Tables.Add(tIva);
        }

        private void FillTab()
        {
            string s = "";

            s = "SELECT * FROM GesDivForTestate ORDER BY dit_num";
            DataTable t = _clsFun.FillTabSql("GesDivForTestate", s, false, _strConSql);

            foreach(DataRow y in t.Rows)
            {
                if ((DateTime)y["dit_sof"] == _clsDef.DAYOUT)
                    y["dit_sof"] = DBNull.Value;
                if ((DateTime)y["dit_soi"] == _clsDef.DAYOUT)
                    y["dit_soi"] = DBNull.Value;
                if ((DateTime)y["dit_sif"] == _clsDef.DAYOUT)
                    y["dit_sif"] = DBNull.Value;
            }

            dgv1.DataSource = t;

            s = "SELECT * FROM GesDivForArticoli ORDER BY div_num";
            t = _clsFun.FillTabSql("GesDivForArticoli", s, false, _strConSql);

            DataColumn[] keys = new DataColumn[4];
            keys[0] = t.Columns["div_tva"];
            keys[1] = t.Columns["div_num"];
            keys[2] = t.Columns["div_arf"];
            keys[3] = t.Columns["div_nri"];
            t.PrimaryKey = keys;
            _dasGen.Tables.Add(t);

            s = "SELECT * FROM GesDivForBarcode ORDER BY die_num";
            t = _clsFun.FillTabSql("GesDivForBarcode", s, false, _strConSql);
            keys = new DataColumn[4];
            keys[0] = t.Columns["die_for"];
            keys[1] = t.Columns["die_num"];
            keys[2] = t.Columns["die_arf"];
            keys[3] = t.Columns["die_ean"];
            t.PrimaryKey = keys;
            _dasGen.Tables.Add(t);

            s = "SELECT * FROM TabIva";
            DataTable tIva = _clsFun.FillTabSql(TABTABIVA, s, false, _strConSql);
            _dasGen.Tables.Add(tIva);
        }

        private void ForImport()
        {
            if (dgv2.DataSource != null)
                ((DataView)dgv2.DataSource).Table.Clear();

            FillRighe();

            Color bc = lblDiv.BackColor;
            lblDiv.BackColor = Color.LightCoral;
            lblDiv.Text = "Caricamento in corso ...";

            DataRow[] j;

            string s = "SELECT * FROM TabForImport WHERE tab_ann=0";
            DataTable t = _clsFun.FillTabSql("TabForImport", s, false, _strConSql);

            foreach (DataRow y in t.Rows)
            {
                if (((string)y["tab_tip"]).Trim() == "TERRON")
                {
                    string sPathFor = ((string)y["tab_pth"]).Trim();

                    string[] sFils = Directory.GetFiles(sPathFor);
                    foreach (string sFil in sFils)
                    {
                        if(Path.GetFileName(sFil).Length >= ((string)y["tab_fil"]).Length)
                        {
                            if (Path.GetFileName(sFil).Substring(0, ((string)y["tab_fil"]).Length) == ((string)y["tab_fil"]).Trim())
                                FillTerronVar(sFil, (string)y["tab_cod"], ((string)y["tab_des"]).Trim());

                            else if (Path.GetFileName(sFil).Substring(0, ((string)y["tab_fat"]).Length) == ((string)y["tab_fat"]).Trim())
                                FillTerronVar(sFil, (string)y["tab_cod"], ((string)y["tab_des"]).Trim());

                            else if (Path.GetFileName(sFil).Substring(0, ((string)y["tab_fio"]).Length) == ((string)y["tab_fio"]).Trim())
                                FillTerronOff(sFil, (string)y["tab_cod"], ((string)y["tab_des"]).Trim());
                        }
                    }
                }
                else if (((string)y["tab_tip"]).Trim() == "BRENDOLAN")
                {
                    string sPathFor = ((string)y["tab_pth"]).Trim();

                    string[] sFils = Directory.GetFiles(sPathFor);
                    foreach (string sFil in sFils)
                    {
                        j = t.Select("tab_fil='" + Path.GetFileName(sFil).Substring(0, 4).ToUpper() + "'");
                        if (j.Length > 0)
                            FillBrendolanArt(sFil, (string)j[0]["tab_cod"], ((string)j[0]["tab_des"]).Trim());

                        j = t.Select("tab_fat='" + Path.GetFileName(sFil).Substring(0, 2).ToUpper() + "'");
                        if (j.Length > 0)
                            FillBrendolanFat(sFil, (string)j[0]["tab_cod"], ((string)j[0]["tab_des"]).Trim());
                    }
                }
                else if (((string)y["tab_tip"]).Trim() == "MIGROSS")
                {
                    string sPathFor = ((string)y["tab_pth"]).Trim();

                    string[] sFils = Directory.GetFiles(sPathFor);
                    foreach (string sFil in sFils)
                    {
                        s = ((string)y["tab_fil"]).Trim();
                        if (Path.GetFileName(sFil).Substring(0, s.Length) == s)
                            FillMigrossArt(sFil, (string)y["tab_cod"], ((string)y["tab_des"]).Trim());

                        s = ((string)y["tab_fat"]).Trim();
                        if (Path.GetFileName(sFil).Substring(0, s.Length) == s)
                            FillMigrossFat(sFil, y);
                    }
                }
                else if (((string)y["tab_tip"]).Trim() == "EUROSPESA")
                {
                    string sPathFor = ((string)y["tab_pth"]).Trim();

                    string[] sFils = Directory.GetFiles(sPathFor);
                    foreach (string sFil in sFils)
                    {
                        //if(sFil.Length )

                        string s2 = "";

                        s = ((string)y["tab_fil"]).Trim();
                        s2 = Path.GetFileName(sFil);
                        if (s != "" && s2.Length >= s.Length && Path.GetFileName(sFil).Substring(0, s.Length).ToUpper() == s)
                            FillEuroSpesaVar(sFil, (string)y["tab_cod"], ((string)y["tab_des"]).Trim());

                        s = ((string)y["tab_fio"]).Trim();
                        s2 = Path.GetFileName(sFil);
                        if (s != "" && s2.Length >= s.Length && Path.GetFileName(sFil).Substring(0, s.Length).ToUpper() == s)
                            FillEuroSpesaOff(sFil, (string)y["tab_cod"], ((string)y["tab_des"]).Trim());

                        s = ((string)y["tab_fat"]).Trim();
                        s2 = Path.GetFileName(sFil);
                        if (s != "" && s2.Length >= s.Length && Path.GetFileName(sFil).Substring(0, s.Length).ToUpper() == s)
                            FillEuroSpesaFat(sFil, (string)y["tab_cod"], ((string)y["tab_des"]).Trim());
                    }
                }
                else if (((string)y["tab_tip"]).Trim() == "DPIU")
                {
                    string sPathFor = ((string)y["tab_pth"]).Trim();

                    string[] sFils = Directory.GetFiles(sPathFor);
                    foreach (string sFil in sFils)
                    {
                        s = ((string)y["tab_fil"]).Trim();

                        string[] a = s.Split(',');

                        foreach (string ss in a)
                        {
                            if (ss != "" && Path.GetFileName(sFil).Substring(0, ss.Length).ToUpper() == ss)
                            {
                                if (Path.GetFileName(sFil).Substring(0, ss.Length).ToUpper() == "VAR")
                                    FillDPiuVar(sFil, (string)y["tab_cod"], ((string)y["tab_des"]).Trim(), "VAR", (string)y["tab_tip"]);
                                else
                                    FillDPiuVar_new(sFil, (string)y["tab_cod"], ((string)y["tab_des"]).Trim(), "VAR", (string)y["tab_tip"]);
                            }
                        }

                        s = ((string)y["tab_fio"]).Trim();
                        if (s != "" && Path.GetFileName(sFil).Substring(0, s.Length).ToUpper() == s)
                            FillDPiuVar(sFil, (string)y["tab_cod"], ((string)y["tab_des"]).Trim(), "OFF", (string)y["tab_tip"]);

                        s = ((string)y["tab_fat"]).Trim();
                        if (s != "" && Path.GetFileName(sFil).Substring(0, s.Length).ToUpper() == s)
                            FillDPiuVar(sFil, (string)y["tab_cod"], ((string)y["tab_des"]).Trim(), "FAT", (string)y["tab_tip"]);
                    }
                }
                else if (((string)y["tab_tip"]).Trim() == "UNICOM")
                {
                    string sPathFor = ((string)y["tab_pth"]).Trim();

                    s = ((string)y["tab_fil"]).Trim();
                    string[] a = s.Split(',');

                    if (a.Length > 0)
                    {
                        string[] sFils = Directory.GetFiles(sPathFor);
                        foreach (string sFil in sFils)
                        {
                            s = a[0];
                            if (s != "" && Path.GetFileName(sFil).Substring(0, s.Length).ToUpper() == s)
                                FillUnicom(sFil, "VAR", y);

                            s = ((string)y["tab_fat"]).Trim();

                            string[] aa = s.Split(',');

                            if (aa.Length > 0)
                            {
                                s = aa[0];
                                if (s != "" && Path.GetFileName(sFil).Substring(0, s.Length).ToUpper() == s)
                                    FillUnicom(sFil, "BOL", y);
                            }
                            if (aa.Length > 1)
                            {
                                s = aa[1];
                                if (s != "" && Path.GetFileName(sFil).Substring(0, s.Length).ToUpper() == s)
                                    FillUnicom(sFil, "FAT", y);
                            }
                        }

                        if(a.Length > 1)
                        {
                            s = a[1];

                            string sPath = Path.GetDirectoryName(s) + "\\";
                            string sFile = Path.GetFileName(s);

                            string[] sFils2 = Directory.GetFiles(sPath);
                            foreach (string sFil in sFils2)
                            {
                                s = Path.GetFileName(sFil);

                                if(s.Length > 10 && s.Substring(0,sFile.Length) == sFile && Path.GetExtension(s) == ".TXT")
                                    FillUnicomPescheria((string)y["tab_cod"], sFil);
                            }
                        }
                    }
                }
                else if (((string)y["tab_tip"]).Trim() == "GOTTARDO")
                {
                    string sPathFor = ((string)y["tab_pth"]).Trim();

                    string[] sFils = Directory.GetFiles(sPathFor);
                    foreach (string sFil in sFils)
                    {
                        s = ((string)y["tab_fil"]).Trim();
                        string[] a = s.Split(',');
                        if (a.Length > 1)
                            s = a[0];
                        if (s != "" && Path.GetFileName(sFil).Length >= s.Length && Path.GetFileName(sFil).Substring(0, s.Length).ToLower() == s && Path.GetExtension(sFil).ToLower() == ".txt")
                            FillGottardo(sFil, "VAR", y); 

                        s = ((string)y["tab_fat"]).Trim();
                        a = s.Split(',');
                        if (a.Length > 1)
                            s = a[0];
                        if (s != "" && Path.GetFileName(sFil).Length >= s.Length && Path.GetFileName(sFil).Substring(0, s.Length).ToUpper() == s && Path.GetExtension(sFil).ToLower() == ".txt")
                            FillGottardo(sFil, "FAT", y); 
                    }
                }
                else if (((string)y["tab_tip"]).Trim() == "ORTOFRU")
                {
                    string sPathFor = ((string)y["tab_pth"]).Trim();

                    string[] sFils = Directory.GetFiles(sPathFor);
                    foreach (string sFil in sFils)
                    {
                        s = ((string)y["tab_fil"]).Trim();

                        string[] a = s.Split(',');

                        foreach (string ss in a)
                        {
                            if (ss != "" && Path.GetFileName(sFil).Substring(0, ss.Length).ToUpper() == ss)
                                FillOrtoFrutticola(sFil, (string)y["tab_cod"], ((string)y["tab_des"]).Trim(), "VAR", (string)y["tab_tip"]);
                        }

                        s = ((string)y["tab_fio"]).Trim();
                        if (s != "" && Path.GetFileName(sFil).Substring(0, s.Length).ToUpper() == s)
                            FillOrtoFrutticola(sFil, (string)y["tab_cod"], ((string)y["tab_des"]).Trim(), "OFF", (string)y["tab_tip"]);

                        s = ((string)y["tab_fat"]).Trim();
                        if (s != "" && Path.GetFileName(sFil).Substring(0, s.Length).ToUpper() == s)
                            FillOrtoFrutticolaFat(sFil, y);
                        //FillOrtoFrutticolaFat(sFil, (string)y["tab_cod"], ((string)y["tab_des"]).Trim());
                    }
                }

                else if (((string)y["tab_tip"]).Trim() == "DMO")
                {
                    string sPathFor = ((string)y["tab_pth"]).Trim();

                    string[] sFils = Directory.GetFiles(sPathFor);
                    foreach (string sFil in sFils)
                    {
                        s = ((string)y["tab_fil"]).Trim();

                        string[] a = s.Split(',');

                        foreach (string ss in a)
                        {
                            if (ss != "" && Path.GetFileName(sFil).Substring(0, ss.Length).ToLower() == ss)
                                FillDmo(sFil, (string)y["tab_cod"], ((string)y["tab_des"]).Trim(), "VAR", (string)y["tab_tip"]);
                        }

                        //s = ((string)y["tab_fio"]).Trim();
                        //if (s != "" && Path.GetFileName(sFil).Substring(0, s.Length).ToUpper() == s)
                        //    FillOrtoFrutticola(sFil, (string)y["tab_cod"], ((string)y["tab_des"]).Trim(), "OFF", (string)y["tab_tip"]);

                        s = ((string)y["tab_fat"]).Trim();
                        if (s != "" && Path.GetFileName(sFil).Substring(0, s.Length).ToUpper() == s)
                            FillDmoFat(sFil, y);
                        //FillDmoFat(sFil, (string)y["tab_cod"], ((string)y["tab_des"]).Trim());
                    }
                }

                else if (((string)y["tab_tip"]).Trim() == "GABBIANO")
                {
                    string sPathFor = ((string)y["tab_pth"]).Trim();

                    string[] sFils = Directory.GetFiles(sPathFor);
                    foreach (string sFil in sFils)
                    {
                        s = ((string)y["tab_fil"]).Trim();

                        string[] a = s.Split(',');

                        foreach (string ss in a)
                        {
                            if (ss != "" && Path.GetFileName(sFil).Substring(0, ss.Length).ToLower() == ss)
                                FillGabbiano(sFil, (string)y["tab_cod"], ((string)y["tab_des"]).Trim(), "VAR", (string)y["tab_tip"]);
                        }

                        //s = ((string)y["tab_fio"]).Trim();
                        //if (s != "" && Path.GetFileName(sFil).Substring(0, s.Length).ToUpper() == s)
                        //    FillOrtoFrutticola(sFil, (string)y["tab_cod"], ((string)y["tab_des"]).Trim(), "OFF", (string)y["tab_tip"]);

                        //s = ((string)y["tab_fat"]).Trim();
                        //if (s != "" && Path.GetFileName(sFil).Substring(0, s.Length).ToUpper() == s)
                        //    FillDmoFat(sFil, (string)y["tab_cod"], ((string)y["tab_des"]).Trim());
                    }
                }

                else if (((string)y["tab_tip"]).Trim() == "FIETTA")
                {
                    string sPathFor = ((string)y["tab_pth"]).Trim();

                    string[] sFils = Directory.GetFiles(sPathFor);
                    foreach (string sFil in sFils)
                    {
                        //s = ((string)y["tab_fil"]).Trim();

                        //string[] a = s.Split(',');

                        //foreach (string ss in a)
                        //{
                        //    if (ss != "" && Path.GetFileName(sFil).Substring(0, ss.Length).ToUpper() == ss)
                        //        FillOrtoFrutticola(sFil, (string)y["tab_cod"], ((string)y["tab_des"]).Trim(), "VAR", (string)y["tab_tip"]);
                        //}

                        //s = ((string)y["tab_fio"]).Trim();
                        //if (s != "" && Path.GetFileName(sFil).Substring(0, s.Length).ToUpper() == s)
                        //    FillOrtoFrutticola(sFil, (string)y["tab_cod"], ((string)y["tab_des"]).Trim(), "OFF", (string)y["tab_tip"]);

                        s = ((string)y["tab_fat"]).Trim();
                        if (s != "" && Path.GetFileName(sFil).Substring(0, s.Length).ToUpper() == s && Path.GetExtension(sFil).ToLower() == ".txt")
                            FillFietta(sFil, "FAT", y);
                        //FillFietta(sFil, (string)y["tab_cod"], ((string)y["tab_des"]).Trim(), "FAT", (string)y["tab_tip"]);
                    }
                }
                else if (false && ((string)y["tab_tip"]).Trim() == "VENFRI")
                {
                    string sPathFor = ((string)y["tab_pth"]).Trim();

                    string[] sFils = Directory.GetFiles(sPathFor);
                    foreach (string sFil in sFils)
                    {
                        //s = ((string)y["tab_fil"]).Trim();

                        //string[] a = s.Split(',');

                        //foreach (string ss in a)
                        //{
                        //    if (ss != "" && Path.GetFileName(sFil).Substring(0, ss.Length).ToUpper() == ss)
                        //        FillOrtoFrutticola(sFil, (string)y["tab_cod"], ((string)y["tab_des"]).Trim(), "VAR", (string)y["tab_tip"]);
                        //}

                        //s = ((string)y["tab_fio"]).Trim();
                        //if (s != "" && Path.GetFileName(sFil).Substring(0, s.Length).ToUpper() == s)
                        //    FillOrtoFrutticola(sFil, (string)y["tab_cod"], ((string)y["tab_des"]).Trim(), "OFF", (string)y["tab_tip"]);

                        s = ((string)y["tab_fat"]).Trim();
                        if (s != "" && Path.GetFileName(sFil).Substring(0, s.Length).ToUpper() == s && Path.GetExtension(sFil).ToLower() == ".txt")
                            FillVefri(sFil, (string)y["tab_cod"], ((string)y["tab_des"]).Trim(), "FAT", (string)y["tab_tip"]);
                    }
                }
                else if (((string)y["tab_tip"]).Trim() == "DADO")
                {
                    string sPathFor = ((string)y["tab_pth"]).Trim();

                    string[] sFils = Directory.GetFiles(sPathFor);
                    foreach (string sFil in sFils)
                    {
                        if (Path.GetExtension(sFil) == ".csv")
                        {
                            j = t.Select("tab_fil='" + Path.GetFileName(sFil).Substring(0, 8).ToUpper() + "'");
                            if (j.Length > 0)
                                FillDadoArt(sFil, (string)j[0]["tab_cod"], ((string)j[0]["tab_des"]).Trim());
                        }
                    }
                }
                else if (((string)y["tab_tip"]).Trim() == "SMA")
                {
                    string sPathFor = ((string)y["tab_pth"]).Trim();

                    string[] sFils = Directory.GetFiles(sPathFor);
                    foreach (string sFil in sFils)
                    {
                        s = (string)y["tab_fil"];

                        if (Path.GetFileName(sFil).Length > s.Length && Path.GetFileName(sFil).Substring(0, s.Length) == s)
                        {
                            FillSMA(sFil, (string)y["tab_cod"], ((string)y["tab_des"]).Trim(), ((string)y["tab_neg"]).Trim());
                        }
                    }
                }
                else if (((string)y["tab_tip"]).Trim() == "ALFA")
                {
                    //DivNewFornitori((string)y["tab_cod"], (string)y["tab_des"]);
                    DivNewFornitori((string)y["tab_tra"], (string)y["tab_cod"], (string)y["tab_des"]);
                }
                else if (((string)y["tab_tip"]).Trim() == "APOFFICE")
                {
                    string sPathFor = ((string)y["tab_pth"]).Trim();

                    string sPathLoca = "";
                    string sTipForDiv = "";

                    s = ((string)y["tab_fil"]).Trim();

                    string[] a = s.Split('|');

                    if (a.Length > 1)
                    {
                        sPathLoca = a[1];

                        s = a[0];
                        string[] aFil = s.Split(',');

                        if (a.Length > 2)
                            sTipForDiv = a[2];

                        //sPathFor = "\\\\192.168.178.39\\scambio\\apsistemi\\";
                        //sPathFor = "D:\\ApProject";

                        if (aFil.Length > 0)
                        {
                            //if( Directory.Exists(sPathFor))

                            //_clsFun.ErrorLog("Divulgazione APOFFICE", "Passo 1 " + sPathFor);

                            if (!_clsFun.DirectoryExists(sPathFor))
                                MessageBox.Show("Collegamento a " + sPathFor + " non riuscito!", "COLLEGAMENTO A SEDE", MessageBoxButtons.OK, MessageBoxIcon.Question);
                            else
                            {
                                //_clsFun.ErrorLog("Divulgazione APOFFICE", "Passo 2");

                                string[] sFils = Directory.GetFiles(sPathFor);
                                foreach (string sFil in sFils)
                                {
                                    //_clsFun.ErrorLog("Divulgazione APOFFICE", "Passo 3");

                                    foreach (string ss in aFil)
                                    {
                                        /*
                                         * Da fare il filtro sul negozio
                                         */ 

                                        if (ss != "" && Path.GetFileName(sFil).Substring(0, ss.Length).ToLower() == ss.ToLower())
                                        {
                                            try
                                            {
                                                //s = sPathLoca + Path.GetFileName(sFil);

                                                s = Path.GetFileName(sFil);

                                                if (s.Length > 50)
                                                    s = s.Substring(0, 50) + ".csv_" + DateTime.Now.ToString("yyyyMMddHHmmss");

                                                string sFii = sPathLoca + ss + DateTime.Now.ToString("yyyyMMdd") + ".csv";

                                                StreamWriter sw = new StreamWriter(sFii, true);

                                                using (StreamReader sr = new StreamReader(sFil, System.Text.Encoding.Default))
                                                {
                                                    string sRig = "";

                                                    while ((sRig = sr.ReadLine()) != null)
                                                    {
                                                        sw.Write(sRig + _clsDef.CRLF);
                                                    }

                                                    sr.Close();
                                                    sr.Dispose();
                                                }
                                                sw.Close();
                                                sw.Dispose();

                                                s = sPathLoca + s;

                                                //File.Move(sFil, s);

                                                s = Path.GetDirectoryName(sFil) + "\\old\\" + Path.GetFileName(sFil) + "_" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".csv";

                                                Console.WriteLine("aaaa");

                                                File.Move(sFil, s);

                                            }
                                            catch (Exception ex)
                                            {
                                                Console.WriteLine(ex.Message);
                                            }

                                        }
                                    }
                                }
                            }

                            s = "";
                            if (s == "")
                            {
                                string[] sFils = Directory.GetFiles(sPathLoca);
                                foreach (string sFil in sFils)
                                {
                                    //s = ((string)y["tab_fil"]).Trim();
                                    //string[] a = s.Split(',');

                                    Boolean b = true;

                                    foreach (string ss in aFil)
                                    {
                                        if (ss != "" && Path.GetFileName(sFil).Substring(0, ss.Length).ToLower() == ss.ToLower())
                                        {
                                            if (ss == "tab_")
                                                new clsVariazioni().AggApOfficeTab(sFil);
                                            //Console.WriteLine("aaaa");
                                            else
                                            {
                                                string sTip = "VAR";
                                                if (ss == "dof_")
                                                    sTip = "OFF";

                                                //FillApOffice(sFil, (string)y["tab_cod"], ((string)y["tab_des"]).Trim(), "VAR", (string)y["tab_tip"]);
                                                FillApOffice(sFil, (string)y["tab_cod"], ((string)y["tab_des"]).Trim(), sTip, (string)y["tab_tip"], sTipForDiv);
                                                b = false;
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
                else if (((string)y["tab_tip"]).Trim() == "APDIV")
                {
                    string sPathFor = ((string)y["tab_pth"]).Trim();

                    string[] sFils = Directory.GetFiles(sPathFor);
                    foreach (string sFil in sFils)
                    {
                        if (Path.GetExtension(sFil) == ".csv")
                        {
                            j = t.Select("tab_fil='" + Path.GetFileName(sFil).Substring(0, 8).ToUpper() + "'");
                            if (j.Length > 0)
                                FillApDivVar(sFil, (string)j[0]["tab_cod"], ((string)j[0]["tab_des"]).Trim());
                        }
                        //j = t.Select("tab_fat='" + Path.GetFileName(sFil).Substring(0, 5).ToUpper() + "'");
                        //if (j.Length > 0)
                        //    FillMigrossFat(sFil, (string)j[0]["tab_cod"], ((string)j[0]["tab_des"]).Trim());

                        //j = t.Select("tab_fio='" + Path.GetFileName(sFil).Substring(0, 7).ToUpper() + "'");
                        //if (j.Length > 0)
                        //    FillTerronOff(sFil, (string)j[0]["tab_cod"], ((string)j[0]["tab_des"]).Trim());
                    }
                }
                else if (((string)y["tab_tip"]).Trim() == "TMPDIV" || ((string)y["tab_tip"]).Trim() == "FTPDIV")
                {
                    //0006 APPHONE(TMPDIV/FTPDIV) ==> TMPDIV

                    FillTmpDiv();
                }
                else if (!DBNull.Value.Equals(y["tab_tra"]) && (string)y["tab_tra"] != "")
                {
                    DivNewFornitori2((string)y["tab_tra"], (string)y["tab_cod"], (string)y["tab_des"]);
                }
            }

            lblCnt.Text = dgv2.RowCount.ToString();

            lblDiv.BackColor = bc;
            lblDiv.Text = "";

            //Controllo documenti se già elaborati
            CtrlDoc();
            //}
        }

        private void CtrlDoc()
        {
            string s = "";
            string sMsg = "";

            foreach (DataRow y in ((DataTable)dgv1.DataSource).Rows)
            {
                if((string)y["dit_tva"] == "FAT")
                {
                    s = "SELECT * FROM GesFatTestate WHERE fat_tpd='FA' AND fat_cfo='" + y["dit_for"] + "' AND fat_ndo='" + y["dit_ndo"] + "' AND fat_ddo=" + _clsFun.DaySql((DateTime)y["dit_ddo"]);
                    DataTable t = _clsFun.FillTabSql("FAT", s, true, _strConSql);
                    if (t.Rows.Count > 0)
                        sMsg += "Fattura " + (string)y["dit_ndo"] + " del " + ((DateTime)y["dit_ddo"]).ToString("dd/MM/yyyy") + _clsDef.CRLF;
                }
            }

            if(sMsg != "")
                MessageBox.Show(sMsg, "DOCUMENTI GIA' PRESENTI", MessageBoxButtons.OK, MessageBoxIcon.Question);

        }

        private void FillTerronVar(string strFil, string strFor, string strDes)
        {
            DataRow x;
            string s = "";
            string sRig = "";

            using (StreamReader sr = new StreamReader(strFil, System.Text.Encoding.Default))
            {
                clsDivTerron clsDiv = new clsDivTerron();

                FileInfo fInfo = new FileInfo(strFil);

                progressBar1.Value = 0;
                progressBar1.Maximum = (int)fInfo.Length;
                progressBar1.Minimum = 0;

                DataTable tTes = (DataTable)dgv1.DataSource;

                DataTable tDie = _dasGen.Tables[TABTMPEAN];
                DataColumn[] keys = new DataColumn[2];
                keys[0] = tDie.Columns["die_arf"];
                keys[1] = tDie.Columns["die_ean"];
                tDie.PrimaryKey = keys;

                DataTable t = _dasGen.Tables[TABTMPFDR];

                string sTva = "";
                string sNdo = "";
                DateTime dDdo = _clsDef.DAYOUT;
                string sNum = "";
                string sNums = "";
                int iRig = 1;
                DataRow[] j;

                sTva = VARART;
                if (Path.GetFileName(strFil).Length > 7 && Path.GetFileName(strFil).Substring(0, 8) == "KONZMOVI")
                    sTva = VARART;
                else if (Path.GetFileName(strFil).Substring(0, 7) == "KONZMOV")
                {
                    if(strFil.ToLower().IndexOf("impianto") < 0)
                        sTva = VARFAT;
                }
                //if (sTva == "F")
                //    sTva = VARFAT;
                //else if (sTva == "A")
                //    sTva = VARART;
                //else if (sTva == "O")
                //    sTva = VAROFF;

                while ((sRig = sr.ReadLine()) != null)
                {
                    progressBar1.Increment(sRig.Length);
                    System.Windows.Forms.Application.DoEvents();

                    //sTva = sRig.Substring(127, 1);

                    if (sRig.Length > 50)
                    {
                        sNdo = sRig.Substring(6, 6);

                        j = tTes.Select("dit_tva='" + sTva + "' AND dit_ndo='" + sNdo + "' AND dit_ddo=#" + dDdo.ToString("MM/dd/yyyy") + "#");
                        if (j.Length == 0)
                        {
                            //sNum = (tTes.Rows.Count + 1).ToString("000");
                            sNum = DitNewNum();

                            sNums += sNum + ";";

                            x = tTes.NewRow();
                            x["dit_num"] = sNum;
                            x["dit_for"] = strFor;
                            x["dit_des"] = strDes;
                            x["dit_fil"] = Path.GetFileName(strFil);
                            x["dit_tva"] = sTva;
                            x["dit_ndo"] = sNdo;
                            x["dit_ddo"] = new DateTime(Convert.ToInt32(sRig.Substring(16, 4)), Convert.ToInt32(sRig.Substring(14, 2)), Convert.ToInt32(sRig.Substring(12, 2)));
                            x["dit_pth"] = Path.GetDirectoryName(strFil);
                            x["dit_nri"] = 1;
                            x["dit_cho"] = "E";
                            tTes.Rows.Add(x);
                            iRig = 1;
                            dDdo = (DateTime)x["dit_ddo"];
                        }
                        else
                        {
                            sNum = (string)j[0]["dit_num"];
                            j[0]["dit_nri"] = (decimal)j[0]["dit_nri"] + 1;
                            iRig = Convert.ToInt32(j[0]["dit_nri"]);
                        }
                        x = t.NewRow();
                        clsDiv.DivArt(sRig, x);     /***/
                        x["div_tva"] = sTva;
                        x["div_num"] = sNum;
                        x["div_nri"] = iRig.ToString("00000");

                        j = t.Select("div_tva='" + sTva + "' AND div_num='" + sNum + "' AND div_arf='" + x["div_arf"] + "'");
                        if (j.Length == 0)
                            t.Rows.Add(x);

                        //if ((string)x["div_arf"] == "026200")
                        //    Console.WriteLine("aaaaaaaaa");

                        s = ((string)x["div_ean"]).Trim();
                        if (s != "")
                        {
                            if (_strPar020EsclusioneEan2021davDiv == "S" && s.Length == 13 && s.Substring(7, 5) == "00000" && ((s.Substring(0, 2) == "20") || s.Substring(0, 2) == "21"))
                                Console.WriteLine("Scarto i barcode a peso che iniziano per 20 e 21");
                            else
                            {
                                s = s.Substring(0, s.Length - 1) + new clsCtrlCodici().FindMod10Digit(s.Substring(0, s.Length - 1));
                                j = tDie.Select("die_arf='" + x["div_arf"] + "' AND die_ean='" + s + "'");
                                if (j.Length == 0)
                                {
                                    DataRow y = tDie.NewRow();
                                    y["die_num"] = sNum;
                                    y["die_for"] = strFor;
                                    y["die_arf"] = x["div_arf"];
                                    y["die_ean"] = s;
                                    tDie.Rows.Add(y);

                                    if (s != ((string)x["div_ean"]).Trim())
                                        _clsFun.ErrorLog("Divulgazione TERRON", "Corretto check digit " + ((string)x["div_ean"]).Trim() + " -> " + s);
                                }
                            }
                        }
                    }
                    //break;
                }

                string[] a = sNums.Split(';');

                for (int i = 0; i < a.Length; i++)
                {
                    DataTable tTab = CtrlDivArt(strFil, t, strFor, a[i]);
                    DataView v = new DataView(tTab, "tab_cod=''", "", DataViewRowState.CurrentRows);
                    if (v.Count > 0)
                    {
                        frmGesDivCtrlTab f = new frmGesDivCtrlTab();
                        f._tabTab = tTab.Copy();
                        f._strForCod = strFor;
                        f.ShowDialog();
                        if (f._bolMdy)
                        {
                            CtrlDivArt(strFil, t, strFor, a[i]);
                        }
                    }
                }
            }
        }

        private void FillTerronOff(string strFil, string strFor, string strDes)
        {
            DataRow x;
            DataRow[] j;
            string s = "";
            string sRig = "";

            using (StreamReader sr = new StreamReader(strFil, System.Text.Encoding.Default))
            {
                clsDivTerron clsDiv = new clsDivTerron();

                FileInfo fInfo = new FileInfo(strFil);

                progressBar1.Value = 0;
                progressBar1.Maximum = (int)fInfo.Length;
                progressBar1.Minimum = 0;

                DataTable tDie = _dasGen.Tables[TABTMPEAN];
                DataTable t = (DataTable)dgv1.DataSource;

                string sNof = "";
                string sNum = "";
                string sNums = "";
                int iRig = 0;

                while ((sRig = sr.ReadLine()) != null)
                {
                    progressBar1.Increment(sRig.Length);
                    System.Windows.Forms.Application.DoEvents();

                    s = sRig.Substring(0, 5);
                    if (s != sNof)
                    {
                        sNof = s;
                        //sNum = (t.Rows.Count + 1).ToString("000");
                        sNum = DitNewNum();

                        sNums += sNum + ";";

                        t = (DataTable)dgv1.DataSource;
                        x = t.NewRow();
                        x["dit_num"] = sNum;
                        x["dit_for"] = strFor;
                        x["dit_des"] = strDes;
                        x["dit_fil"] = Path.GetFileName(strFil);
                        x["dit_tva"] = VAROFF;
                        x["dit_ndo"] = sNof;
                        x["dit_pth"] = Path.GetDirectoryName(strFil);
                        //x["dit_ddo"] = new DateTime(Convert.ToInt32(sRig.Substring(16, 4)), Convert.ToInt32(sRig.Substring(14, 2)), Convert.ToInt32(sRig.Substring(12, 2)));
                        x["dit_sii"] = _clsFun.Str2Day(sRig.Substring(65, 8));
                        x["dit_sif"] = _clsFun.Str2Day(sRig.Substring(73, 8));
                        x["dit_soi"] = _clsFun.Str2Day(sRig.Substring(81, 8));
                        x["dit_sof"] = _clsFun.Str2Day(sRig.Substring(89, 8));
                        x["dit_cho"] = "E";
                        t.Rows.Add(x);

                        t = _dasGen.Tables[TABTMPFDR];
                        iRig = 0;
                    }

                    iRig++;
                    x = t.NewRow();
                    clsDiv.DivOff(sRig, x);
                    if ((string)x["div_arf"] == "240220" || (string)x["div_arf"] == "792075")
                        Console.WriteLine("aaaaaaaa");

                    x["div_tva"] = VAROFF;
                    x["div_num"] = sNum;
                    x["div_nri"] = iRig.ToString("00000");
                    t.Rows.Add(x);

                    /* Barcode */

                    s = "";
                    if(!DBNull.Value.Equals(x["div_ean"]))
                        s = ((string)x["div_ean"]).Trim();
                    if (s != "")
                    {
                        if (_strPar020EsclusioneEan2021davDiv == "S" && s.Length == 13 && s.Substring(7, 5) == "00000" && ((s.Substring(0, 2) == "20") || s.Substring(0, 2) == "21"))
                            Console.WriteLine("Scarto i barcode a peso che iniziano per 20 e 21");
                        else
                        {
                            s = s.Substring(0, s.Length - 1) + new clsCtrlCodici().FindMod10Digit(s.Substring(0, s.Length - 1));
                            j = tDie.Select("die_arf='" + x["div_arf"] + "' AND die_ean='" + s + "'");
                            if (j.Length == 0)
                            {
                                DataRow y = tDie.NewRow();
                                y["die_num"] = sNum;
                                y["die_for"] = strFor;
                                y["die_arf"] = x["div_arf"];
                                y["die_ean"] = s;
                                tDie.Rows.Add(y);

                                if (s != ((string)x["div_ean"]).Trim())
                                    _clsFun.ErrorLog("Divulgazione TERRON", "Corretto check digit " + ((string)x["div_ean"]).Trim() + " -> " + s);
                            }
                        }
                    }

                }
                
                string[] a = sNums.Split(';');

                for (int i = 0; i < a.Length; i++)
                    CtrlDivArt(strFil, t, strFor, a[i]);

                //s = "";
                //foreach(DataRow y in t.Rows)
                //{
                //    if(((string)y["div_art"]).Trim() == "")
                //    {
                //        s += (string)y["div_arf"] + " " + (string)y["div_ard"] + _clsDef.CRLF;
                //    }
                //}
                //if (s != "")
                //{
                //    MessageBox.Show(s, "ARTICOLI NON TROVATI (scartati)");
                //}
            }
        }

        private void FillBrendolanFat(string strFil, string strFor, string strDes)
        {
            DataRow x;
            string s = "";
            string sRig = "";

            using (StreamReader sr = new StreamReader(strFil, System.Text.Encoding.Default))
            {
                clsDivBrendolan clsDiv = new clsDivBrendolan();

                FileInfo fInfo = new FileInfo(strFil);

                progressBar1.Value = 0;
                progressBar1.Maximum = (int)fInfo.Length;
                progressBar1.Minimum = 0;

                DataTable tTes = (DataTable)dgv1.DataSource;

                DataTable tDie = _dasGen.Tables[TABTMPEAN];
                DataColumn[] keys = new DataColumn[2];
                keys[0] = tDie.Columns["die_arf"];
                keys[1] = tDie.Columns["die_ean"];
                tDie.PrimaryKey = keys;

                DataTable t = _dasGen.Tables[TABTMPFDR];

                string sTva = "FAT";
                string sTri = "";
                string sNdo = "";
                DateTime dDdo = _clsDef.DAYOUT;
                string sNum = "";
                string sNums = "";
                int iRig = 1;
                DataRow[] j;

                while ((sRig = sr.ReadLine()) != null)
                {
                    progressBar1.Increment(sRig.Length);
                    System.Windows.Forms.Application.DoEvents();

                    if (sRig.Length > 0)
                    {
                        sTri = sRig.Substring(0, 1);
                        sNdo = sRig.Substring(1, 7);

                        if (sTri == "A")
                        {
                            //s = sRig.Substring(8, 6);
                            //dDdo = new DateTime(Convert.ToInt32("20" + s.Substring(0, 2)), Convert.ToInt32(s.Substring(2, 2)), Convert.ToInt32(s.Substring(4, 2)));
                            dDdo = _clsFun.Str2Day("20" + sRig.Substring(8, 6));
                            j = tTes.Select("dit_tva='" + sTva + "' AND dit_ndo='" + sNdo + "' AND dit_ddo=#" + dDdo.ToString("MM/dd/yyyy") + "#");
                            if (j.Length == 0)
                            {
                                //sNum = (tTes.Rows.Count + 1).ToString("000");
                                sNum = DitNewNum();

                                sNums += sNum + ";";

                                x = tTes.NewRow();
                                x["dit_num"] = sNum;
                                x["dit_for"] = strFor;
                                x["dit_des"] = strDes;
                                x["dit_fil"] = Path.GetFileName(strFil);
                                x["dit_tva"] = sTva;
                                x["dit_ndo"] = sNdo;
                                x["dit_ddo"] = dDdo;
                                x["dit_pth"] = Path.GetDirectoryName(strFil);
                                x["dit_nri"] = 1;
                                x["dit_cho"] = "E";

                                tTes.Rows.Add(x);
                                iRig = 0;
                                dDdo = (DateTime)x["dit_ddo"];
                            }
                        }
                        //}
                        //    sNum = (string)j[0]["dit_num"];
                        //    j[0]["dit_nri"] = (decimal)j[0]["dit_nri"] + 1;
                        //    iRig = Convert.ToInt32(j[0]["dit_nri"]);
                        //}
                        else if (sTri == "B")
                        {
                            x = t.NewRow();
                            clsDiv.DivFat(sRig, x);     /***/
                            x["div_tva"] = sTva;
                            x["div_num"] = sNum;
                            x["div_dva"] = dDdo;
                            x["div_nri"] = iRig.ToString("00000");

                            j = t.Select("div_tva='" + sTva + "' AND div_num='" + sNum + "' AND div_arf='" + x["div_arf"] + "'");
                            if (j.Length == 0)
                                t.Rows.Add(x);

                            //if ((string)x["div_arf"] == "026200")
                            //    Console.WriteLine("aaaaaaaaa");

                            s = ((string)x["div_ean"]).Trim();
                            if (s != "" && s.Length > 7 && _clsFun.Numerico(s))
                            {
                                if (_strPar020EsclusioneEan2021davDiv == "S" && s.Length == 13 && s.Substring(7, 5) == "00000" && ((s.Substring(0, 2) == "20") || s.Substring(0, 2) == "21"))
                                    Console.WriteLine("Scarto i barcode a peso che iniziano per 20 e 21");
                                else
                                {
                                    s = s.Substring(0, s.Length - 1) + new clsCtrlCodici().FindMod10Digit(s.Substring(0, s.Length - 1));
                                    j = tDie.Select("die_arf='" + x["div_arf"] + "' AND die_ean='" + s + "'");
                                    if (j.Length == 0)
                                    {
                                        DataRow y = tDie.NewRow();
                                        y["die_num"] = sNum;
                                        y["die_for"] = strFor;
                                        y["die_arf"] = x["div_arf"];
                                        y["die_ean"] = s;
                                        tDie.Rows.Add(y);

                                        if (s != ((string)x["div_ean"]).Trim())
                                            _clsFun.ErrorLog("Divulgazione BRENDOLAN", "Corretto check digit " + ((string)x["div_ean"]).Trim() + " -> " + s);
                                    }
                                }
                            }
                        }
                    }
                    //break;
                    iRig++;
                }

                string[] a = sNums.Split(';');
                for (int i = 0; i < a.Length; i++)
                {
                    DataTable tTab = CtrlDivArt(strFil, t, strFor, a[i]);
                    DataView v = new DataView(tTab, "tab_cod=''", "", DataViewRowState.CurrentRows);
                    if (v.Count > 0)
                    {
                        frmGesDivCtrlTab f = new frmGesDivCtrlTab();
                        f._tabTab = tTab.Copy();
                        f._strForCod = strFor;
                        f.ShowDialog();
                        if (f._bolMdy)
                        {
                            CtrlDivArt(strFil, t, strFor, a[i]);
                        }
                    }
                }
            }
        }

        private void FillBrendolanArt(string strFil, string strFor, string strDes)
        {
            DataRow x;
            string s = "";
            string sRig = "";

            using (StreamReader sr = new StreamReader(strFil, System.Text.Encoding.Default))
            {
                clsDivBrendolan clsDiv = new clsDivBrendolan();

                FileInfo fInfo = new FileInfo(strFil);

                progressBar1.Value = 0;
                progressBar1.Maximum = (int)fInfo.Length;
                progressBar1.Minimum = 0;

                DataTable tTes = (DataTable)dgv1.DataSource;

                DataTable tDie = _dasGen.Tables[TABTMPEAN];
                DataColumn[] keys = new DataColumn[2];
                keys[0] = tDie.Columns["die_arf"];
                keys[1] = tDie.Columns["die_ean"];
                tDie.PrimaryKey = keys;

                DataTable t = _dasGen.Tables[TABTMPFDR];

                string sTva = "ART";
                string sTri = "";
                string sNdo = "";
                DateTime dDdo = _clsDef.DAYOUT;
                string sNum = "";
                string sNums = "";
                int iRig = 1;
                DataRow[] j;

                while ((sRig = sr.ReadLine()) != null)
                {
                    progressBar1.Increment(sRig.Length);
                    System.Windows.Forms.Application.DoEvents();

                    if (sRig.Length > 396)
                    {
                        if (iRig == 1)
                        {
                            s = sRig.Substring(209, 8);
                            try
                            {
                                dDdo = new DateTime(Convert.ToInt32(s.Substring(0, 4)), Convert.ToInt32(s.Substring(4, 2)), Convert.ToInt32(s.Substring(6, 2)));
                            }
                            catch (Exception ex)
                            {
                                dDdo = DateTime.Today;
                            }

                            j = tTes.Select("dit_tva='" + sTva + "' AND dit_ndo='" + sNdo + "' AND dit_ddo=#" + dDdo.ToString("MM/dd/yyyy") + "#");
                            if (j.Length == 0)
                            {
                                //sNum = (tTes.Rows.Count + 1).ToString("000");
                                sNum = DitNewNum();

                                sNums += sNum + ";";
                                x = tTes.NewRow();
                                x["dit_num"] = sNum;
                                x["dit_for"] = strFor;
                                x["dit_des"] = strDes;
                                x["dit_fil"] = Path.GetFileName(strFil);
                                x["dit_tva"] = sTva;
                                x["dit_ndo"] = sNdo;
                                x["dit_ddo"] = dDdo;
                                x["dit_pth"] = Path.GetDirectoryName(strFil);
                                x["dit_nri"] = 1;
                                x["dit_cho"] = "E";

                                tTes.Rows.Add(x);
                                iRig = 1;
                                dDdo = (DateTime)x["dit_ddo"];
                            }
                        }

                        Boolean b = true;
                        x = t.NewRow();
                        clsDiv.DivArt(sRig, x);     /***/
                        x["div_tva"] = sTva;
                        x["div_num"] = sNum;
                        x["div_nri"] = iRig.ToString("000000");

                        //Seck 20200423 Controllo codice fornitore
                        s = _clsFun.CtrlCrt(Convert.ToString(x["div_arf"]), "1234567890qwertyuiopasdfghjklzxcvbnmQWERTYUIOPASDFGHJKLZXCVBNM");

                        if (s == "987786")
                            Console.WriteLine("xxxxxxx");

                        if (s.Trim() != Convert.ToString(x["div_arf"]).Trim())
                            b = false;

                        //_clsFun.ErrorLog("Controllo articolo", "Arf " + Convert.ToString(x["div_arf"]));

                        if (b)
                        {
                            j = t.Select("div_tva='" + sTva + "' AND div_num='" + sNum + "' AND div_arf='" + x["div_arf"] + "'");
                            if (j.Length == 0)
                                t.Rows.Add(x);

                            s = ((string)x["div_ean"]).Trim();
                            if (s != "" && s.Length > 7 && _clsFun.Numerico(s))
                            {
                                if (_strPar020EsclusioneEan2021davDiv == "S" && s.Length == 13 && s.Substring(7, 5) == "00000" && ((s.Substring(0, 2) == "20") || s.Substring(0, 2) == "21"))
                                    Console.WriteLine("Scarto i barcode a peso che iniziano per 20 e 21");
                                else
                                {
                                    s = s.Substring(0, s.Length - 1) + new clsCtrlCodici().FindMod10Digit(s.Substring(0, s.Length - 1));
                                    j = tDie.Select("die_arf='" + x["div_arf"] + "' AND die_ean='" + s + "'");
                                    if (j.Length == 0)
                                    {
                                        DataRow y = tDie.NewRow();
                                        y["die_num"] = sNum;
                                        y["die_for"] = strFor;
                                        y["die_arf"] = x["div_arf"];
                                        y["die_ean"] = s;
                                        tDie.Rows.Add(y);

                                        if (s != ((string)x["div_ean"]).Trim())
                                            _clsFun.ErrorLog("Divulgazione BRENDOLAN", "Corretto check digit " + ((string)x["div_ean"]).Trim() + " -> " + s);
                                    }
                                }
                            }
                        }
                    }

                    iRig++;
                    //}
                    //break;

                    //lblDiv.Text = iRig.ToString();
                }

                string[] a = sNums.Split(';');
                for (int i = 0; i < a.Length; i++)
                {
                    s = a[i];
                    //if (s == "" || s.Length <= 7 || !_clsFun.Numerico(s))
                    //    Console.WriteLine("Scarto");
                    //else if (_strPar020EsclusioneEan2021davDiv == "S" && s.Length == 13 && s.Substring(7, 5) == "00000" && ((s.Substring(0, 2) == "20") || s.Substring(0, 2) == "21"))
                    //    Console.WriteLine("Scarto i barcode a peso che iniziano per 20 e 21");
                    //else
                    if (s != "")
                    {
                        DataTable tTab = CtrlDivArt(strFil, t, strFor, a[i]);
                        DataView v = new DataView(tTab, "tab_cod=''", "", DataViewRowState.CurrentRows);
                        if (v.Count > 0)
                        {
                            frmGesDivCtrlTab f = new frmGesDivCtrlTab();
                            f._tabTab = tTab.Copy();
                            f._tabDiv = t.Copy();
                            f._strForCod = strFor;
                            f.ShowDialog();
                            if (f._bolMdy)
                            {
                                CtrlDivArt(strFil, t, strFor, a[i]);
                            }
                        }
                    }
                }
            }
        }

        private void FillMigrossArt(string strFil, string strFor, string strDes)
        {
            DataRow x;
            string s = "";
            string sRig = "";

            using (StreamReader sr = new StreamReader(strFil, System.Text.Encoding.Default))
            {
                clsDivMigross clsDiv = new clsDivMigross();
                clsDiv._strPar017EcrDivFornitori = _strPar017EcrDivFornitori;

                FileInfo fInfo = new FileInfo(strFil);

                progressBar1.Value = 0;
                progressBar1.Maximum = (int)fInfo.Length;
                progressBar1.Minimum = 0;

                DataTable tTes = (DataTable)dgv1.DataSource;

                DataTable tDie = _dasGen.Tables[TABTMPEAN];
                DataColumn[] keys = new DataColumn[2];
                keys[0] = tDie.Columns["die_arf"];
                keys[1] = tDie.Columns["die_ean"];
                tDie.PrimaryKey = keys;

                DataTable t = _dasGen.Tables[TABTMPFDR];

                string sTva = "ART";
                string sNdo = "";
                DateTime dDdo = _clsDef.DAYOUT;
                DateTime dDti = _clsDef.DAYOUT;
                DateTime dDtf = _clsDef.DAYOUT;
                string sNum = "";
                string sNums = "";
                int iRig = 1;
                DataRow[] j;
                Boolean b = true;

                while ((sRig = sr.ReadLine()) != null)
                {
                    progressBar1.Increment(sRig.Length);
                    System.Windows.Forms.Application.DoEvents();

                    sTva = "";
                    if (sRig.Length > 10 && sRig.Substring(0, 2) == "71")
                        sTva = "ART";
                    if (sRig.Length > 10 && sRig.Substring(0, 2) == "73")
                        sTva = VAROFF;
                    if(sTva != "")
                    {
                        if (sTva == "ART")
                        {
                            b = false;
                            s = sRig.Substring(184, 8);
                            dDdo = _clsFun.Str2Day(sRig.Substring(184, 8)); 
                            j = tTes.Select("dit_tva='" + sTva + "' AND dit_ndo='" + sNdo + "' AND dit_ddo=#" + dDdo.ToString("MM/dd/yyyy") + "#");
                            if (j.Length == 0)
                            {
                                //sNum = (tTes.Rows.Count + 1).ToString("000");
                                sNum = DitNewNum();

                                sNums += sNum + ";";
                                x = tTes.NewRow();
                                x["dit_num"] = sNum;
                                x["dit_for"] = strFor;
                                x["dit_des"] = strDes;
                                x["dit_fil"] = Path.GetFileName(strFil);
                                x["dit_tva"] = sTva;
                                x["dit_ndo"] = sNdo;
                                x["dit_ddo"] = dDdo;
                                x["dit_pth"] = Path.GetDirectoryName(strFil);
                                x["dit_nri"] = 1;
                                x["dit_cho"] = "E";

                                tTes.Rows.Add(x);
                                iRig = 1;
                                dDdo = (DateTime)x["dit_ddo"];
                            }
                        }

                        if (sTva == VAROFF)
                        {
                            b = false;
                            s = sRig.Substring(184, 8);
                            sNdo = sRig.Substring(131, 5);
                            dDdo = _clsFun.Str2Day(sRig.Substring(184, 8));
                            dDti = _clsFun.Str2Day(sRig.Substring(192, 8));
                            dDtf = _clsFun.Str2Day(sRig.Substring(200, 8));
                            j = tTes.Select("dit_tva='" + sTva + "' AND dit_soi=#" + dDti.ToString("MM/dd/yyyy") + "# AND dit_sof=#" + dDtf.ToString("MM/dd/yyyy") + "#");
                            if (j.Length == 0)
                            {
                                //sNum = (tTes.Rows.Count + 1).ToString("000");
                                sNum = DitNewNum();
                                sNums += sNum + ";";

                                x = tTes.NewRow();
                                //x["dit_num"] = sNum;
                                //x["dit_for"] = strFor;
                                //x["dit_des"] = strDes;
                                //x["dit_fil"] = Path.GetFileName(strFil);
                                //x["dit_tva"] = sTva;
                                //x["dit_ndo"] = sNdo;
                                //x["dit_ddo"] = dDdo;
                                //x["dit_pth"] = Path.GetDirectoryName(strFil);
                                //x["dit_nri"] = 1;
                                //x["dit_soi"] = dDti;
                                //x["dit_sof"] = dDtf;

                                x["dit_num"] = sNum;
                                x["dit_for"] = strFor;
                                x["dit_des"] = strDes;
                                x["dit_fil"] = Path.GetFileName(strFil);
                                x["dit_tva"] = VAROFF;
                                x["dit_ndo"] = sNdo;
                                x["dit_pth"] = Path.GetDirectoryName(strFil);
                                //x["dit_ddo"] = new DateTime(Convert.ToInt32(sRig.Substring(16, 4)), Convert.ToInt32(sRig.Substring(14, 2)), Convert.ToInt32(sRig.Substring(12, 2)));
                                x["dit_sii"] = dDti;
                                x["dit_sif"] = dDtf;
                                x["dit_soi"] = dDti;
                                x["dit_sof"] = dDtf;
                                x["dit_cho"] = "E";

                                tTes.Rows.Add(x);
                                iRig = 1;
                                //dDdo = (DateTime)x["dit_ddo"];
                            }
                        }

                        x = t.NewRow();
                        clsDiv.DivArt(sRig, x);     /***/
                        x["div_tva"] = sTva;
                        x["div_num"] = sNum;
                        x["div_nri"] = iRig.ToString("00000");

                        j = t.Select("div_tva='" + sTva + "' AND div_num='" + sNum + "' AND div_arf='" + x["div_arf"] + "'");
                        if (j.Length == 0)
                            t.Rows.Add(x);

                        if ((string)x["div_arf"] == "026200")
                            Console.WriteLine("aaaaaaaaa");

                        s = "";
                        if (!DBNull.Value.Equals(x["div_ean"]))
                            s = ((string)x["div_ean"]).Trim();

                        if (_strPar020EsclusioneEan2021davDiv == "S" && s.Length == 13 && s.Substring(7, 5) == "00000" && ((s.Substring(0, 2) == "20") || s.Substring(0, 2) == "21"))
                            Console.WriteLine("Scarto i barcode a peso che iniziano per 20 e 21");
                        else if (s != "" && s.Length > 7 && _clsFun.Numerico(s))
                        {
                            s = s.Substring(0, s.Length - 1) + new clsCtrlCodici().FindMod10Digit(s.Substring(0, s.Length - 1));
                            j = tDie.Select("die_arf='" + x["div_arf"] + "' AND die_ean='" + s + "'");
                            if (j.Length == 0)
                            {
                                DataRow y = tDie.NewRow();
                                y["die_num"] = sNum;
                                y["die_for"] = strFor;
                                y["die_arf"] = x["div_arf"];
                                y["die_ean"] = s;
                                tDie.Rows.Add(y);

                                if (s != ((string)x["div_ean"]).Trim())
                                    _clsFun.ErrorLog("Divulgazione MIGROSS", "Corretto check digit " + ((string)x["div_ean"]).Trim() + " -> " + s);
                            }
                        }
                    }

                    iRig++;
                    //}
                    //break;
                }

                string[] a = sNums.Split(';');
                for (int i = 0; i < a.Length; i++)
                {
                    if (a[i].Trim() != "")
                    {
                        DataTable tTab = CtrlDivArt(strFil, t, strFor, a[i]);
                        DataView v = new DataView(tTab, "tab_cod=''", "", DataViewRowState.CurrentRows);
                        if (v.Count > 0)
                        {
                            frmGesDivCtrlTab f = new frmGesDivCtrlTab();
                            f._tabTab = tTab.Copy();
                            f._tabDiv = t.Copy();
                            f._tabDie = tDie.Copy();
                            f._strForCod = strFor;
                            f.ShowDialog();
                            if (f._bolMdy)
                            {
                                CtrlDivArt(strFil, t, strFor, a[i]);
                            }
                        }
                    }
                }
            }
        }

        //private void FillMigrossFat(string strFil, string strFor, string strDes, string strNeg)
        private void FillMigrossFat(string strFil, DataRow rowFat)
        {
            DataRow x;
            string s = "";
            string sRig = "";

            string strFor = (string)rowFat["tab_cod"];
            string strDes = (string)rowFat["tab_des"];
            string strNeg = (string)rowFat["tab_neg"];

            //(string)j[0]["tab_cod"], ((string)j[0]["tab_des"]).Trim(), ((string)j[0]["tab_neg"]).Trim()

            using (StreamReader sr = new StreamReader(strFil, System.Text.Encoding.Default))
            {
                clsDivMigross clsDiv = new clsDivMigross();

                FileInfo fInfo = new FileInfo(strFil);

                progressBar1.Value = 0;
                progressBar1.Maximum = (int)fInfo.Length;
                progressBar1.Minimum = 0;

                DataTable tTes = (DataTable)dgv1.DataSource;

                DataTable tDie = _dasGen.Tables[TABTMPEAN];
                DataColumn[] keys = new DataColumn[2];
                keys[0] = tDie.Columns["die_arf"];
                keys[1] = tDie.Columns["die_ean"];
                tDie.PrimaryKey = keys;

                DataTable t = _dasGen.Tables[TABTMPFDR];

                string sTva = "FAT";
                string sNdo = "";
                DateTime dDdo = _clsDef.DAYOUT;
                string sNum = "";
                string sNums = "";
                int iRig = 0;
                DataRow[] j;

                while ((sRig = sr.ReadLine()) != null)
                {
                    progressBar1.Increment(sRig.Length);
                    System.Windows.Forms.Application.DoEvents();

                    if (sRig.Length > 10)
                    {
                        sNdo = sRig.Substring(11, 7);
                        dDdo = _clsFun.Str2Day(sRig.Substring(20, 8));
                        j = tTes.Select("dit_tva='" + sTva + "' AND dit_ndo='" + sNdo + "' AND dit_ddo=#" + dDdo.ToString("MM/dd/yyyy") + "#");
                        if (j.Length == 0)
                        {
                            string sNeg = "001";

                            string[] aNeg = strNeg.Split(',');
                            if (aNeg.Length > 1)
                            {
                                s = sRig.Substring(0, 6);
                                for (int ii = 0; ii < aNeg.Length; ii++)
                                {
                                    if (aNeg[ii] == s)
                                    {
                                        sNeg = (ii + 1).ToString("000");
                                        break;
                                    }
                                }
                            }

                            //sNum = (tTes.Rows.Count + 1).ToString("000");
                            sNum = DitNewNum();

                            sNums += sNum + ";";
                            x = tTes.NewRow();
                            x["dit_num"] = sNum;
                            x["dit_neg"] = sNeg;
                            x["dit_for"] = strFor;
                            x["dit_des"] = strDes;
                            x["dit_fil"] = Path.GetFileName(strFil);
                            x["dit_tva"] = sTva;
                            x["dit_ndo"] = sNdo;
                            x["dit_ddo"] = dDdo;
                            x["dit_pth"] = Path.GetDirectoryName(strFil);
                            x["dit_nri"] = 1;
                            x["dit_cho"] = "E";

                            tTes.Rows.Add(x);
                            iRig = 0;
                            dDdo = (DateTime)x["dit_ddo"];
                        }
                        iRig++;
                        x = t.NewRow();
                        clsDiv.DivFat(sRig, x);     /***/
                        x["div_tva"] = sTva;
                        x["div_num"] = sNum;
                        x["div_dva"] = dDdo;
                        x["div_nri"] = iRig.ToString("00000");

                        j = t.Select("div_tva='" + sTva + "' AND div_num='" + sNum + "' AND div_arf='" + x["div_arf"] + "'");
                        if (j.Length == 0)
                            t.Rows.Add(x);

                        //if ((string)x["div_arf"] == "026200")
                        //    Console.WriteLine("aaaaaaaaa");

                        s = ((string)x["div_ean"]).Trim();
                        if (s != "" && s.Length > 7 && _clsFun.Numerico(s))
                        {
                            if (_strPar020EsclusioneEan2021davDiv == "S" && s.Length == 13 && s.Substring(7, 5) == "00000" && ((s.Substring(0, 2) == "20") || s.Substring(0, 2) == "21"))
                                Console.WriteLine("Scarto i barcode a peso che iniziano per 20 e 21");
                            else
                            {
                                s = s.Substring(0, s.Length - 1) + new clsCtrlCodici().FindMod10Digit(s.Substring(0, s.Length - 1));
                                j = tDie.Select("die_arf='" + x["div_arf"] + "' AND die_ean='" + s + "'");
                                if (j.Length == 0)
                                {
                                    DataRow y = tDie.NewRow();
                                    y["die_num"] = sNum;
                                    y["die_for"] = strFor;
                                    y["die_arf"] = x["div_arf"];
                                    y["die_ean"] = s;
                                    tDie.Rows.Add(y);

                                    if (s != ((string)x["div_ean"]).Trim())
                                        _clsFun.ErrorLog("Divulgazione MIGROSS", "Corretto check digit " + ((string)x["div_ean"]).Trim() + " -> " + s);
                                }
                            }
                        }
                    }
                }

                string[] a = sNums.Split(';');
                for (int i = 0; i < a.Length; i++)
                {
                    if (a[i].Trim() != "")
                    {

                        DataTable tTab = CtrlDivArt(strFil, t, strFor, a[i]);
                        DataView v = new DataView(tTab, "tab_cod=''", "", DataViewRowState.CurrentRows);
                        if (v.Count > 0)
                        {
                            frmGesDivCtrlTab f = new frmGesDivCtrlTab();
                            f._tabTab = tTab.Copy();
                            f._strForCod = strFor;
                            f._tabDie = tDie.Copy();
                            f.ShowDialog();
                            if (f._bolMdy)
                            {
                                CtrlDivArt(strFil, t, strFor, a[i]);
                            }
                        }
                    }
                }
            }
        }

        private void FillEuroSpesaVar(string strFil, string strFor, string strDes)
        {
            DataRow x;
            string s = "";
            string sRig = "";

            using (StreamReader sr = new StreamReader(strFil, System.Text.Encoding.Default))
            {
                clsDivEuroSpesa clsDiv = new clsDivEuroSpesa();

                FileInfo fInfo = new FileInfo(strFil);

                progressBar1.Value = 0;
                progressBar1.Maximum = (int)fInfo.Length;
                progressBar1.Minimum = 0;

                DataTable tTes = (DataTable)dgv1.DataSource;

                DataTable tDie = _dasGen.Tables[TABTMPEAN];
                DataColumn[] keys = new DataColumn[2];
                keys[0] = tDie.Columns["die_arf"];
                keys[1] = tDie.Columns["die_ean"];
                tDie.PrimaryKey = keys;

                DataTable t = _dasGen.Tables[TABTMPFDR];

                string sTva = "";
                //string sNdo = "";
                DateTime dDdo = _clsDef.DAYOUT;
                string sNum = "";
                string sNums = "";
                int iRig = 1;
                DataRow[] j;

                sTva = VARART;
                //if (Path.GetFileName(strFil).Substring(0, 7) == "KONZMOV")
                //{
                //    if (strFil.ToLower().IndexOf("impianto") < 0)
                //        sTva = VARFAT;
                //}

                //sNum = (tTes.Rows.Count + 1).ToString("000");
                sNum = DitNewNum();

                while ((sRig = sr.ReadLine()) != null)
                {
                    progressBar1.Increment(sRig.Length);
                    System.Windows.Forms.Application.DoEvents();

                    //Boolean b = false;

                    if(sRig.Length > 20)
                    {
                        string sTre = sRig.Substring(0, 3);

                        //if (sTre == "ART")
                        //{
                        //    b = true;
                        //    sNdo = sRig.Substring(6, 6);
                        //}


                        if ("ART-ARE-ARB-ARF".Contains(sTre))
                        {

                            //j = tTes.Select("dit_tva='" + sTva + "' AND dit_ndo='" + sNdo + "' AND dit_ddo=#" + dDdo.ToString("MM/dd/yyyy") + "#");
                            j = tTes.Select("dit_tva='" + sTva + "' AND dit_ndo='" + sNum + "' AND dit_ddo=#" + dDdo.ToString("MM/dd/yyyy") + "#");
                            if (j.Length == 0)
                            {
                                sNums += sNum + ";";
                                //sNdo = sNum;

                                x = tTes.NewRow();
                                x["dit_num"] = sNum;
                                x["dit_for"] = strFor;
                                x["dit_des"] = strDes;
                                x["dit_fil"] = Path.GetFileName(strFil);
                                x["dit_tva"] = sTva;
                                x["dit_ndo"] = sNum;    //sNdo
                                //x["dit_ddo"] = new DateTime(Convert.ToInt32(sRig.Substring(16, 4)), Convert.ToInt32(sRig.Substring(14, 2)), Convert.ToInt32(sRig.Substring(12, 2)));
                                x["dit_ddo"] = DateTime.Today;
                                x["dit_pth"] = Path.GetDirectoryName(strFil);
                                x["dit_nri"] = 1;
                                x["dit_cho"] = "E";

                                tTes.Rows.Add(x);
                                iRig = 1;
                                dDdo = (DateTime)x["dit_ddo"];
                            }
                            else
                            {
                                sNum = (string)j[0]["dit_num"];
                                j[0]["dit_nri"] = (decimal)j[0]["dit_nri"] + 1;
                                iRig = Convert.ToInt32(j[0]["dit_nri"]);
                            }
                            x = t.NewRow();
                            if (sTre == "ART")
                            {
                                clsDiv.DivArt(sRig, x);     /* Lettura record */
                                x["div_tva"] = sTva;
                                x["div_num"] = sNum;
                                x["div_nri"] = iRig.ToString("00000");

                                j = t.Select("div_tva='" + sTva + "' AND div_num='" + sNum + "' AND div_arf='" + x["div_arf"] + "'");
                                if (j.Length == 0)
                                    t.Rows.Add(x);
                            }

                            if (sTre == "ARE")
                            {
                                string sArf = sRig.Substring(5, 15).Trim();
                                string sEan = sRig.Substring(20, 13).Trim();

                                if (sEan != "")
                                {
                                    if (_strPar020EsclusioneEan2021davDiv == "S" && s.Length == 13 && s.Substring(7, 5) == "00000" && ((s.Substring(0, 2) == "20") || s.Substring(0, 2) == "21"))
                                        Console.WriteLine("Scarto i barcode a peso che iniziano per 20 e 21");
                                    else
                                    {
                                        if (sEan.Length == 7 && sEan.Substring(0, 2) == "20")
                                        {
                                            s = sEan + "000000";
                                            sEan = s;
                                        }
                                        else if (sEan.Length == 13 && sEan.Substring(0, 1) == "2")
                                            Console.WriteLine("Faccio niente");
                                        else if (sEan.Length > 7)
                                        {
                                            s = sEan.Substring(0, sEan.Length - 1) + new clsCtrlCodici().FindMod10Digit(sEan.Substring(0, sEan.Length - 1));
                                            if (s != sEan)
                                                _clsFun.ErrorLog("Divulgazione EUROSPESA", "Corretto check digit " + sEan + " -> " + s);
                                        }
                                        else
                                            s = sEan;

                                        j = tDie.Select("die_arf='" + sArf + "' AND die_ean='" + s + "'");
                                        if (j.Length == 0)
                                        {
                                            DataRow y = tDie.NewRow();
                                            y["die_num"] = sNum;
                                            y["die_for"] = strFor;
                                            y["die_arf"] = sArf;
                                            y["die_ean"] = s;
                                            tDie.Rows.Add(y);
                                        }
                                    }
                                }
                            }

                            if (sTre == "ARB")
                            {
                                string sArf = sRig.Substring(5, 15).Trim();
                                //string sEan = sRig.Substring(20, 13).Trim();

                                j = t.Select("div_tva='" + sTva + "' AND div_num='" + sNum + "' AND div_arf='" + sArf + "'");
                                if (j.Length > 0)
                                {
                                    clsDiv.DivBil(sRig, j[0]);
                                    Console.WriteLine("aaaaaaa");
                                }
                            }
                            if (sTre == "ARF")
                            {
                                string sArf = sRig.Substring(5, 15).Trim();

                                j = t.Select("div_tva='" + sTva + "' AND div_num='" + sNum + "' AND div_arf='" + sArf + "'");
                                if (j.Length > 0)
                                {
                                    clsDiv.DivArf(sRig, j[0]);
                                    Console.WriteLine("aaaaaaa");
                                }
                            }
                        }
                    }
                    //break;
                }

                string[] a = sNums.Split(';');

                for (int i = 0; i < a.Length; i++)
                {
                    DataTable tTab = CtrlDivArt(strFil, t, strFor, a[i]);
                    DataView v = new DataView(tTab, "tab_cod=''", "", DataViewRowState.CurrentRows);
                    if (v.Count > 0)
                    {
                        frmGesDivCtrlTab f = new frmGesDivCtrlTab();
                        f._tabTab = tTab.Copy();
                        f._tabDiv = t.Copy();
                        f._strForCod = strFor;
                        f.ShowDialog();
                        if (f._bolMdy)
                        {
                            CtrlDivArt(strFil, t, strFor, a[i]);
                        }
                    }
                }
            }
        }

        private void FillEuroSpesaOff(string strFil, string strFor, string strDes)
        {
            DataRow x;
            DataRow[] j;
            string s = "";
            string sRig = "";

            using (StreamReader sr = new StreamReader(strFil, System.Text.Encoding.Default))
            {
                clsDivEuroSpesa clsDiv = new clsDivEuroSpesa();

                FileInfo fInfo = new FileInfo(strFil);

                progressBar1.Value = 0;
                progressBar1.Maximum = (int)fInfo.Length;
                progressBar1.Minimum = 0;

                DataTable tDie = _dasGen.Tables[TABTMPEAN];
                DataTable tTes = (DataTable)dgv1.DataSource;
                DataTable t = new DataTable();

                string sNof = "";
                string sNum = "";
                string sNums = "";
                int iRig = 0;

                while ((sRig = sr.ReadLine()) != null)
                {
                    progressBar1.Increment(sRig.Length);
                    System.Windows.Forms.Application.DoEvents();

                    if(sRig.Length > 40)
                    {
                        sNof = sRig.Substring(7, 3);

                        if (sRig.Substring(0, 1) == "1")
                        {
                            tTes = (DataTable)dgv1.DataSource;
                            //sNum = (tTes.Rows.Count + 1).ToString("000");
                            sNum = DitNewNum();

                            sNums += sNum + ";";

                            x = tTes.NewRow();
                            x["dit_num"] = sNum;
                            x["dit_for"] = strFor;
                            x["dit_des"] = sRig.Substring(10,35).Trim();
                            x["dit_fil"] = Path.GetFileName(strFil);
                            x["dit_tva"] = VAROFF;
                            x["dit_ndo"] = sNof;
                            x["dit_pth"] = Path.GetDirectoryName(strFil);

                            s = sRig.Substring(61, 8);
                            s = "20" + s.Substring(6, 2) + s.Substring(3, 2) + s.Substring(0, 2);
                            x["dit_sii"] = _clsFun.Str2Day(s);

                            s = sRig.Substring(69, 8);
                            s = "20" + s.Substring(6, 2) + s.Substring(3, 2) + s.Substring(0, 2);
                            x["dit_sif"] = _clsFun.Str2Day(s);

                            s = sRig.Substring(45, 8);
                            s = "20" + s.Substring(6, 2) + s.Substring(3, 2) + s.Substring(0, 2);
                            x["dit_soi"] = _clsFun.Str2Day(s);

                            s = sRig.Substring(53, 8);
                            s = "20" + s.Substring(6, 2) + s.Substring(3, 2) + s.Substring(0, 2);
                            x["dit_sof"] = _clsFun.Str2Day(s);

                            x["dit_cho"] = "E";

                            tTes.Rows.Add(x);

                            t = _dasGen.Tables[TABTMPFDR];
                            iRig = 0;
                        }
                        else
                        {
                            iRig++;
                            x = t.NewRow();
                            clsDiv.DivOff(sRig, x);
                            if ((string)x["div_arf"] == "240220" || (string)x["div_arf"] == "792075")
                                Console.WriteLine("aaaaaaaa");

                            x["div_tva"] = VAROFF;
                            x["div_num"] = sNum;
                            x["div_nri"] = iRig.ToString("00000");
                            t.Rows.Add(x);

                            /* Barcode */

                            s = "";
                            if (!DBNull.Value.Equals(x["div_ean"]))
                                s = ((string)x["div_ean"]).Trim();
                            if (s != "")
                            {
                                if (_strPar020EsclusioneEan2021davDiv == "S" && s.Length == 13 && s.Substring(7, 5) == "00000" && ((s.Substring(0, 2) == "20") || s.Substring(0, 2) == "21"))
                                    Console.WriteLine("Scarto i barcode a peso che iniziano per 20 e 21");
                                else
                                {
                                    s = s.Substring(0, s.Length - 1) + new clsCtrlCodici().FindMod10Digit(s.Substring(0, s.Length - 1));
                                    j = tDie.Select("die_arf='" + x["div_arf"] + "' AND die_ean='" + s + "'");
                                    if (j.Length == 0)
                                    {
                                        DataRow y = tDie.NewRow();
                                        y["die_num"] = sNum;
                                        y["die_for"] = strFor;
                                        y["die_arf"] = x["div_arf"];
                                        y["die_ean"] = s;
                                        tDie.Rows.Add(y);

                                        if (s != ((string)x["div_ean"]).Trim())
                                            _clsFun.ErrorLog("Divulgazione TERRON", "Corretto check digit " + ((string)x["div_ean"]).Trim() + " -> " + s);
                                    }
                                }
                            }
                        }
                    }
                }

                string[] a = sNums.Split(';');

                if (t.Rows.Count > 0)
                {
                    for (int i = 0; i < a.Length; i++)
                        CtrlDivArt(strFil, t, strFor, a[i]);
                }
                //s = "";
                //foreach(DataRow y in t.Rows)
                //{
                //    if(((string)y["div_art"]).Trim() == "")
                //    {
                //        s += (string)y["div_arf"] + " " + (string)y["div_ard"] + _clsDef.CRLF;
                //    }
                //}
                //if (s != "")
                //{
                //    MessageBox.Show(s, "ARTICOLI NON TROVATI (scartati)");
                //}
            }
        }

        private void FillEuroSpesaFat(string strFil, string strFor, string strDes)
        {
            DataRow x;
            string s = "";
            string sRig = "";

            using (StreamReader sr = new StreamReader(strFil, System.Text.Encoding.Default))
            {
                clsDivEuroSpesa clsDiv = new clsDivEuroSpesa();

                FileInfo fInfo = new FileInfo(strFil);

                progressBar1.Value = 0;
                progressBar1.Maximum = (int)fInfo.Length;
                progressBar1.Minimum = 0;

                DataTable tTes = (DataTable)dgv1.DataSource;

                //DataTable tDie = _dasGen.Tables[TABTMPEAN];
                //DataColumn[] keys = new DataColumn[2];
                //keys[0] = tDie.Columns["die_arf"];
                //keys[1] = tDie.Columns["die_ean"];
                //tDie.PrimaryKey = keys;

                DataTable t = _dasGen.Tables[TABTMPFDR];

                string sTva = "FAT";
                string sTri = "";
                string sNdo = "";
                DateTime dDdo = _clsDef.DAYOUT;
                string sNum = "";
                string sNums = "";
                int iRig = 1;
                DataRow[] j;

                while ((sRig = sr.ReadLine()) != null)
                {
                    progressBar1.Increment(sRig.Length);
                    System.Windows.Forms.Application.DoEvents();

                    if (sRig.Length > 2)
                    {
                        sTri = sRig.Substring(0, 2);
                        sNdo = sRig.Substring(8, 6);

                        if (sTri == "EN")
                        {
                            dDdo = _clsFun.Str2Day("20" + sRig.Substring(14, 6));
                            j = tTes.Select("dit_tva='" + sTva + "' AND dit_ndo='" + sNdo + "' AND dit_ddo=#" + dDdo.ToString("MM/dd/yyyy") + "#");
                            if (j.Length == 0)
                            {
                                //sNum = (tTes.Rows.Count + 1).ToString("000");
                                sNum = DitNewNum();
                                sNums += sNum + ";";

                                x = tTes.NewRow();
                                x["dit_num"] = sNum;
                                x["dit_for"] = strFor;
                                x["dit_des"] = strDes;
                                x["dit_fil"] = Path.GetFileName(strFil);
                                x["dit_tva"] = sTva;
                                x["dit_ndo"] = sNdo;
                                x["dit_ddo"] = dDdo;
                                x["dit_pth"] = Path.GetDirectoryName(strFil);
                                x["dit_nri"] = 1;
                                x["dit_cho"] = "E";

                                tTes.Rows.Add(x);
                                iRig = 0;
                                dDdo = (DateTime)x["dit_ddo"];
                            }

                            x = t.NewRow();
                            clsDiv.DivFat(sRig, x);     /***/
                            x["div_tva"] = sTva;
                            x["div_num"] = sNum;
                            x["div_dva"] = dDdo;
                            x["div_nri"] = iRig.ToString("00000");

                            j = t.Select("div_tva='" + sTva + "' AND div_num='" + sNum + "' AND div_arf='" + x["div_arf"] + "'");
                            if (j.Length == 0)
                                t.Rows.Add(x);

                            //if ((string)x["div_arf"] == "026200")
                            //    Console.WriteLine("aaaaaaaaa");

                            //s = ((string)x["div_ean"]).Trim();
                            //if (s != "" && s.Length > 7 && _clsFun.Numerico(s))
                            //{
                            //    s = s.Substring(0, s.Length - 1) + new clsCtrlCodici().FindMod10Digit(s.Substring(0, s.Length - 1));
                            //    j = tDie.Select("die_arf='" + x["div_arf"] + "' AND die_ean='" + s + "'");
                            //    if (j.Length == 0)
                            //    {
                            //        DataRow y = tDie.NewRow();
                            //        y["die_for"] = strFor;
                            //        y["die_arf"] = x["div_arf"];
                            //        y["die_ean"] = s;
                            //        tDie.Rows.Add(y);

                            //        if (s != ((string)x["div_ean"]).Trim())
                            //            _clsFun.ErrorLog("Divulgazione BRENDOLAN", "Corretto check digit " + ((string)x["div_ean"]).Trim() + " -> " + s);
                            //    }
                            //}
                        }
                    }
                    //break;
                    iRig++;

                }

                string[] a = sNums.Split(';');
                for (int i = 0; i < a.Length; i++)
                {

                    DataTable tTab = CtrlDivArt(strFil, t, strFor, a[i]);
                    DataView v = new DataView(tTab, "tab_cod=''", "", DataViewRowState.CurrentRows);
                    if (v.Count > 0)
                    {
                        frmGesDivCtrlTab f = new frmGesDivCtrlTab();
                        f._tabTab = tTab.Copy();
                        f._strForCod = strFor;
                        f.ShowDialog();
                        if (f._bolMdy)
                        {
                            CtrlDivArt(strFil, t, strFor, a[i]);
                        }
                    }
                }
            }
        }

        /**** DPIU inizio ***/

        private void FillDPiuVar(string strFil, string strFor, string strDes, string strTip, string strTfo)
        {
            DataRow x;
            string s = "";
            string sRig = "";

            using (StreamReader sr = new StreamReader(strFil, System.Text.Encoding.Default))
            {
                string sTip = strTip;
                if(Path.GetFileName(strFil).Substring(0,1).ToUpper() == "L")
                    sTip = "LIS";

                clsDivDPiu clsDiv = new clsDivDPiu();

                FileInfo fInfo = new FileInfo(strFil);

                progressBar1.Value = 0;
                progressBar1.Maximum = (int)fInfo.Length;
                progressBar1.Minimum = 0;

                DataTable tTes = (DataTable)dgv1.DataSource;

                DataTable tDie = _dasGen.Tables[TABTMPEAN];
                DataColumn[] keys = new DataColumn[2];
                keys[0] = tDie.Columns["die_arf"];
                keys[1] = tDie.Columns["die_ean"];
                tDie.PrimaryKey = keys;

                DataTable t = _dasGen.Tables[TABTMPFDR];

                string sTva = "";
                string sNdo = "";
                DateTime dDdo = _clsDef.DAYOUT;
                string sNum = "";
                string sNums = "";
                int iRig = 1;
                DataRow[] j;

                sTva = VARART;
                if (sTip == "OFF")
                    sTva = "OFF";
                else if (sTip == "FAT")
                    sTva = "FAT";

                //sNum = (tTes.Rows.Count + 1).ToString("000");
                sNum = DitNewNum();

                while ((sRig = sr.ReadLine()) != null)
                {
                    progressBar1.Increment(sRig.Length);
                    System.Windows.Forms.Application.DoEvents();

                    if (sRig.Length > 10)
                    {
                        //sNum = (tTes.Rows.Count).ToString("000");
                        //sNum = DitNewNum();
                        sNum = "001";
                        if (Convert.ToDecimal(tTes.Rows.Count) > 0)
                            sNum = (string)tTes.Rows[tTes.Rows.Count - 1]["dit_num"];

                        sNdo = sNum;
                        if (sTip == "OFF")
                            sNdo = sRig.Substring(22, 5);
                        else if (sTip == "FAT")
                            sNdo = sRig.Substring(1, 7);

                        s = sRig.Substring(66, 14).Trim();
                        if (s == "0001134")
                            Console.WriteLine("aaaa");
                        if (s == "0417493")
                            Console.WriteLine("aaaa");

                        j = tTes.Select("dit_tva='" + sTva + "' AND dit_ndo='" + sNdo + "' AND dit_ddo=#" + dDdo.ToString("MM/dd/yyyy") + "#");

                        if (j.Length == 0)
                        {
                            //sNum = (tTes.Rows.Count + 1).ToString("000");
                            sNum = DitNewNum();

                            sNums += sNum + ";";

                            x = tTes.NewRow();
                            x["dit_num"] = sNum;
                            x["dit_for"] = strFor;
                            x["dit_des"] = strDes;
                            x["dit_fil"] = Path.GetFileName(strFil);
                            x["dit_tva"] = sTva;

                            x["dit_ndo"] = sNum;    //sNdo
                            if (sTip == "OFF")
                                x["dit_ndo"] = sNdo;
                            else if (sTip == "FAT")
                                x["dit_ndo"] = sNdo;

                            x["dit_ddo"] = DateTime.Today;
                            if (sTva == "FAT")
                                x["dit_ddo"] = DateTime.Today;

                            x["dit_pth"] = Path.GetDirectoryName(strFil);
                            x["dit_nri"] = 1;
                            x["dit_tfo"] = strTfo;
                            x["dit_cho"] = "E";

                            if (sTip == "OFF")
                            {
                                x["dit_ndo"] = sNdo;

                                //s = sRig.Substring(61, 8);
                                //s = "20" + s.Substring(6, 2) + s.Substring(3, 2) + s.Substring(0, 2);
                                //x["dit_sii"] = _clsFun.Str2Day(s);

                                //s = sRig.Substring(6, 8);
                                //s = "20" + s.Substring(6, 2) + s.Substring(3, 2) + s.Substring(0, 2);
                                //x["dit_sif"] = _clsFun.Str2Day(s);

                                s = sRig.Substring(6, 8);
                                //s = "20" + s.Substring(6, 2) + s.Substring(3, 2) + s.Substring(0, 2);
                                x["dit_soi"] = _clsFun.Str2Day(s);

                                s = sRig.Substring(14, 8);
                                //s = "20" + s.Substring(6, 2) + s.Substring(3, 2) + s.Substring(0, 2);
                                x["dit_sof"] = _clsFun.Str2Day(s);

                                x["dit_des"] = sRig.Substring(27, 30).Trim();

                            }

                            tTes.Rows.Add(x);
                            iRig = 1;
                            dDdo = (DateTime)x["dit_ddo"];
                        }
                        else
                        {
                            sNum = (string)j[0]["dit_num"];
                        }

                        x = t.NewRow();
                        if (sTip == "LIS")
                            clsDiv.DivLis(sRig, x);     /* Lettura record */
                        else if (sTip == "VAR")
                                clsDiv.DivArt(sRig, x);         /* Lettura record */
                        else if (sTip == "OFF")
                            clsDiv.DivOff(sRig, x);     /* Lettura record */
                        else if (sTip == "FAT")
                            clsDiv.DivFat(sRig, x);     /* Lettura record */
                        x["div_tva"] = sTva;
                        x["div_num"] = sNum;
                        x["div_nri"] = iRig.ToString("00000");
                        iRig++;

                        j = t.Select("div_tva='" + sTva + "' AND div_num='" + sNum + "' AND div_arf='" + x["div_arf"] + "'");
                        if (j.Length == 0)
                            t.Rows.Add(x);

                        /* Barcode */
                        //if (s != "")

                        if (DBNull.Value.Equals(x["div_ean"]))
                            x["div_ean"] = "";

                        if (((string)x["div_ean"]).Trim() != "")
                        {
                            s = ((string)x["div_ean"]).Trim();
                            if (s.Length > 7)
                            {
                                if (_strPar020EsclusioneEan2021davDiv == "S" && s.Length == 13 && s.Substring(7, 5) == "00000" && ((s.Substring(0, 2) == "20") || s.Substring(0, 2) == "21"))
                                    Console.WriteLine("Scarto i barcode a peso che iniziano per 20 e 21");
                                else
                                {
                                    s = s.Substring(0, s.Length - 1) + new clsCtrlCodici().FindMod10Digit(s.Substring(0, s.Length - 1));
                                    j = tDie.Select("die_arf='" + x["div_arf"] + "' AND die_ean='" + s + "'");
                                    if (j.Length == 0)
                                    {
                                        DataRow y = tDie.NewRow();
                                        y["die_num"] = sNum;
                                        y["die_for"] = strFor;
                                        y["die_arf"] = x["div_arf"];
                                        y["die_ean"] = s;
                                        tDie.Rows.Add(y);

                                        if (s != ((string)x["div_ean"]).Trim())
                                            _clsFun.ErrorLog("Divulgazione DPIU", "Corretto check digit " + ((string)x["div_ean"]).Trim() + " -> " + s);
                                    }
                                }
                            }
                        }
                    }
                }

                string[] a = sNums.Split(';');

                for (int i = 0; i < a.Length; i++)
                {
                    DataTable tTab = CtrlDivArt(strFil, t, strFor, a[i]);
                    if (sTip != "OFF")          //Evito controllo articoli
                    {
                        DataView v = new DataView(tTab, "tab_cod=''", "", DataViewRowState.CurrentRows);
                        if (v.Count > 0)
                        {
                            frmGesDivCtrlTab f = new frmGesDivCtrlTab();
                            f._tabTab = tTab.Copy();
                            f._tabDiv = t.Copy();
                            f._strForCod = strFor;
                            f.ShowDialog();
                            if (f._bolMdy)
                            {
                                CtrlDivArt(strFil, t, strFor, a[i]);
                            }
                        }
                    }
                }
            }
        }

        private void FillDPiuVar_new(string strFil, string strFor, string strDes, string strTip, string strTfo)
        {
            DataRow x;
            string s = "";
            //string sRig = "";
            clsDivDPiu clsDiv = new clsDivDPiu();

            lblDiv.Text = "Caricamento listino DPiU'";
            DataTable tVar = clsDiv.Fil2Tab(strFil, progressBar1);

            progressBar1.Value = 0;
            progressBar1.Maximum = tVar.Rows.Count;
            progressBar1.Minimum = 0;

            //using (StreamReader sr = new StreamReader(strFil, System.Text.Encoding.Default))

            if(tVar.Rows.Count > 0)
            {
                progressBar1.Increment(1);
                System.Windows.Forms.Application.DoEvents();

                //string sTip = strTip;
                //if (Path.GetFileName(strFil).Substring(0, 1).ToUpper() == "L")
                //    sTip = "LIS";
                //FileInfo fInfo = new FileInfo(strFil);

                DataTable tTes = (DataTable)dgv1.DataSource;

                DataTable tDie = _dasGen.Tables[TABTMPEAN];
                DataColumn[] keys = new DataColumn[2];
                keys[0] = tDie.Columns["die_arf"];
                keys[1] = tDie.Columns["die_ean"];
                tDie.PrimaryKey = keys;

                DataTable t = _dasGen.Tables[TABTMPFDR];

                string sTva = "";
                string sTip = "";
                string sNdo = "";
                DateTime dDdo = _clsDef.DAYOUT;
                string sNum = "";
                string sNums = "";
                int iRig = 1;
                DataRow[] j;

                sTva = VARART;
                sTip = "VAR";

                //sNum = (tTes.Rows.Count + 1).ToString("000");
                sNum = DitNewNum();

                //while ((sRig = sr.ReadLine()) != null)
                foreach (DataRow y in tVar.Rows)
                {
                    sNum = "001";
                    if (Convert.ToDecimal(tTes.Rows.Count) > 0)
                        sNum = (string)tTes.Rows[tTes.Rows.Count - 1]["dit_num"];

                    //sNdo = sNum;
                    //if (sTip == "OFF")
                    //    sNdo = ""; //sRig.Substring(22, 5);
                    //else if (sTip == "FAT")
                    //    sNdo = ""; // sRig.Substring(1, 7);

                    if((string)y["art_tva"] == "FAT")
                    {
                        sTip = "FAT";
                        sTva = "FAT";
                        if ((string)y["fat_tdo"] == "N")
                            sTva = VARNCA;
                        sNdo = (string)y["fat_ndo"];
                        dDdo = _clsFun.Str2Day((string)y["fat_ddo"]);
                    }
                    else if ((string)y["art_tva"] == "ART")
                    {
                        sTip = "VAR";
                        sTva = "ART";
                        sNdo = "";
                        dDdo = DateTime.Today;
                    }
                    else if ((string)y["art_tva"] == "OFF")
                    {
                        sTip = "OFF";
                        sTva = "OFF";
                        sNdo = (string)y["fat_ndo"];
                        dDdo = DateTime.Today;
                    }

                    j = tTes.Select("dit_tva='" + sTva + "' AND dit_ndo='" + sNdo + "' AND dit_ddo=#" + dDdo.ToString("MM/dd/yyyy") + "#");
                    if (j.Length == 0)
                    {
                        sNum = DitNewNum();

                        sNums += sNum + ";";

                        x = tTes.NewRow();
                        x["dit_num"] = sNum;
                        x["dit_for"] = strFor;
                        x["dit_des"] = strDes;
                        x["dit_fil"] = Path.GetFileName(strFil);
                        x["dit_tva"] = sTva;

                        x["dit_ndo"] = "";          //sNdo
                        if (sTip == "FAT")
                            x["dit_ndo"] = sNum;    //sNdo
                        if (sTip == "OFF")
                            x["dit_ndo"] = sNdo;
                        else if (sTip == "FAT")
                            x["dit_ndo"] = sNdo;

                        x["dit_ddo"] = DateTime.Today;
                        if (sTva == "FAT" || sTva == VARNCA)
                            x["dit_ddo"] = dDdo;

                        x["dit_pth"] = Path.GetDirectoryName(strFil);
                        x["dit_nri"] = 1;
                        x["dit_tfo"] = strTfo;
                        x["dit_cho"] = "E";

                        if (sTip == "OFF")
                        {
                            x["dit_ndo"] = (string)y["fat_ndo"];

                            s = (string)y["off_day"];

                            string[] aa = s.Split('-');

                            x["dit_soi"] = _clsFun.Str2Day(aa[1]);
                            x["dit_sof"] = _clsFun.Str2Day(aa[2]);

                            //s = sRig.Substring(61, 8);
                            //s = "20" + s.Substring(6, 2) + s.Substring(3, 2) + s.Substring(0, 2);
                            //x["dit_sii"] = _clsFun.Str2Day(s);

                            //s = sRig.Substring(6, 8);
                            //s = "20" + s.Substring(6, 2) + s.Substring(3, 2) + s.Substring(0, 2);
                            //x["dit_sif"] = _clsFun.Str2Day(s);

                            //s = ""; // sRig.Substring(6, 8);
                            //x["dit_soi"] = _clsFun.Str2Day(s);

                            ////s = sRig.Substring(14, 8);
                            ////s = "20" + s.Substring(6, 2) + s.Substring(3, 2) + s.Substring(0, 2);
                            //x["dit_sof"] = _clsFun.Str2Day(s);
                            x["dit_des"] = ""; // sRig.Substring(27, 30).Trim();
                        }

                        tTes.Rows.Add(x);
                        iRig = 1;
                        dDdo = (DateTime)x["dit_ddo"];
                    }
                    else
                    {
                        sNum = (string)j[0]["dit_num"];
                    }

                    if (sTva == VARNCA)
                        Console.WriteLine("xyz");

                    x = t.NewRow();
                    //if (sTip == "LIS")
                    //    clsDiv.DivLis(sRig, x);     /* Lettura record */
                    //else if (sTip == "VAR")

                    string sEan = "";

                    if (sTip == "VAR")
                    {

                        x["div_dva"] = DateTime.Today;
                        x["div_arf"] = y["art_arf"];
                        x["div_ard"] = y["art_ard"];
                        x["for_tgr"] = y["art_tgr"];
                        x["div_pne"] = y["art_pne"];
                        x["for_ecr"] = "099099099"; // y["art_ecr"];
                        x["for_umi"] = y["art_umi"];
                        x["div_stf"] = y["art_stf"];
                        x["for_iva"] = y["art_iva"];
                        x["div_adb"] = y["art_adb"];
                        x["div_pxc"] = y["art_pxc"];
                        x["div_cos"] = y["art_pco"];

                        x["div_cos"] = 0;
                        if (!DBNull.Value.Equals(y["art_pco"]))
                            x["div_cos"] = y["art_pco"];

                        x["div_prp"] = y["art_pve"];
                        x["for_rep"] = y["art_rep"];
                        sEan = ((string)y["tmp_ean"]).Trim();
                    }
                    else if (sTip == "FAT")
                    {
                        x["div_dva"] = DateTime.Today;
                        x["div_arf"] = y["art_arf"];
                        x["div_ard"] = y["art_ard"];
                        x["for_tgr"] = y["art_tgr"];
                        x["div_pne"] = y["art_pne"];
                        x["for_ecr"] = "099099099"; // y["art_ecr"];
                        x["for_umi"] = y["art_umi"];
                        x["div_stf"] = y["art_stf"];
                        x["for_iva"] = y["art_iva"];
                        x["div_adb"] = y["art_adb"];
                        x["div_pxc"] = y["art_pxc"];

                        x["div_cos"] = 0;
                        if (!DBNull.Value.Equals(y["art_pco"]))
                            x["div_cos"] = y["art_pco"];

                        x["div_prp"] = y["art_pve"];
                        x["for_rep"] = y["art_rep"];
                        sEan = ((string)y["tmp_ean"]).Trim();

                        //Data
                        //x["div_dva"] = _clsFun.Str2Day(s);

                        //Codice articolo                                
                        //x["div_arf"] = strRig.Substring(63, 7).Trim();

                        //Descriz. articolo              
                        //rowArt["div_ard"] = strRig.Substring(70, 35).Trim();

                        //Quantità fattura               
                        //s = strRig.Substring(105, 9);
                        //if (_clsFun.Numerico(s))
                        //    rowArt["div_qta"] = Convert.ToDecimal(s) / 100;
                        //else
                        x["div_qta"] = (decimal)y["fat_qta"];

                        //Importo
                        //decimal d = Convert.ToDecimal(strRig.Substring(115, 11)) / 1000;
                        x["div_imp"] = (decimal)y["fat_imp"];

                        //IVA
                        //x["for_iva"] = (string)y["fat_iva"];

                        //Tipo documento
                        x["div_tva"] = "FAT";

                        //Importo unitario         
                        x["div_cos"] = (decimal)y["fat_cos"];
                    }
                    else if (sTip == "OFF")
                    {

                        x["div_dva"] = DateTime.Today;
                        x["div_arf"] = y["art_arf"];
                        x["div_ard"] = y["art_ard"];
                        x["for_tgr"] = y["art_tgr"];
                        x["div_pne"] = y["art_pne"];
                        x["for_ecr"] = "099099099"; // y["art_ecr"];
                        x["for_umi"] = y["art_umi"];
                        x["div_stf"] = y["art_stf"];
                        x["for_iva"] = y["art_iva"];
                        x["div_adb"] = y["art_adb"];
                        x["div_pxc"] = y["art_pxc"];
                        x["div_cos"] = y["art_pco"];
                        x["div_prp"] = y["art_pve"];
                        x["for_rep"] = y["art_rep"];
                        sEan = ((string)y["tmp_ean"]).Trim();

                        //Data
                        //x["div_dva"] = _clsFun.Str2Day(s);

                        //Codice articolo                                
                        //x["div_arf"] = strRig.Substring(63, 7).Trim();

                        //Descriz. articolo              
                        //rowArt["div_ard"] = strRig.Substring(70, 35).Trim();

                        //Quantità fattura               
                        //s = strRig.Substring(105, 9);
                        //if (_clsFun.Numerico(s))
                        //    rowArt["div_qta"] = Convert.ToDecimal(s) / 100;
                        //else
                        x["div_qta"] = (decimal)y["fat_qta"];

                        //Importo
                        //decimal d = Convert.ToDecimal(strRig.Substring(115, 11)) / 1000;
                        x["div_imp"] = (decimal)y["fat_imp"];

                        //IVA
                        //x["for_iva"] = (string)y["fat_iva"];

                        //Tipo documento
                        x["div_tva"] = "FAT";

                        //Importo unitario         
                        x["div_cos"] = (decimal)y["fat_cos"];
                    }

                    //clsDiv.DivArt(sRig, x);         /* Lettura record */
                    //else if (sTip == "OFF")
                    //    clsDiv.DivOff(sRig, x);     /* Lettura record */
                    //else if (sTip == "FAT")
                    //    clsDiv.DivFat(sRig, x);     /* Lettura record */
                    x["div_tva"] = sTva;
                    x["div_num"] = sNum;
                    x["div_nri"] = iRig.ToString("00000");
                    iRig++;

                    j = t.Select("div_tva='" + sTva + "' AND div_num='" + sNum + "' AND div_arf='" + x["div_arf"] + "'");
                    if (j.Length == 0)
                        t.Rows.Add(x);
                    else
                    {
                        if ((string)x["div_arf"] == "6872403")
                            Console.WriteLine("XYZ");

                        if (DBNull.Value.Equals(j[0]["div_qta"]))
                            j[0]["div_qta"] = 0;
                        if (!DBNull.Value.Equals(y["fat_qta"]))
                            j[0]["div_qta"] = (decimal)j[0]["div_qta"] + (decimal)y["fat_qta"];
                    }

                    /* Barcode */
                    //if (s != "")

                    //if (DBNull.Value.Equals(x["div_ean"]))
                    //    x["div_ean"] = "";

                    if (sEan != "")
                    {
                        //s = ((string)x["div_ean"]).Trim();
                        string[] aa = sEan.Split(';');

                        foreach (string ss in aa)
                        {
                            //s = ((string)x["div_ean"]).Trim();
                            s = ss;
                            if (s.Length > 7)
                            {
                                if (_strPar020EsclusioneEan2021davDiv == "S" && s.Length == 13 && s.Substring(7, 5) == "00000" && ((s.Substring(0, 2) == "20") || s.Substring(0, 2) == "21"))
                                    Console.WriteLine("Scarto i barcode a peso che iniziano per 20 e 21");
                                else
                                {
                                    s = s.Substring(0, s.Length - 1) + new clsCtrlCodici().FindMod10Digit(s.Substring(0, s.Length - 1));
                                    j = tDie.Select("die_arf='" + x["div_arf"] + "' AND die_ean='" + s + "'");
                                    if (j.Length == 0)
                                    {
                                        DataRow yy = tDie.NewRow();
                                        yy["die_num"] = sNum;
                                        yy["die_for"] = strFor;
                                        yy["die_arf"] = x["div_arf"];
                                        yy["die_ean"] = s;
                                        tDie.Rows.Add(yy);

                                        //if (s != ((string)x["div_ean"]).Trim())
                                        if (s != ss)
                                            _clsFun.ErrorLog("Divulgazione DPIU", "Corretto check digit " + ss + " -> " + s);
                                    }
                                }
                            }
                        }
                    }
                    //}
                }

                string[] a = sNums.Split(';');

                for (int i = 0; i < a.Length; i++)
                {
                    DataTable tTab = CtrlDivArt(strFil, t, strFor, a[i]);
                    if (sTip != "OFF")          //Evito controllo articoli
                    {
                        DataView v = new DataView(tTab, "tab_cod=''", "", DataViewRowState.CurrentRows);
                        if (v.Count > 0)
                        {
                            frmGesDivCtrlTab f = new frmGesDivCtrlTab();
                            f._tabTab = tTab.Copy();
                            f._tabDiv = t.Copy();
                            f._strForCod = strFor;
                            f.ShowDialog();
                            if (f._bolMdy)
                            {
                                CtrlDivArt(strFil, t, strFor, a[i]);
                            }
                        }
                    }
                }
            }
        }

        /*** DPIU fine ***/

        /**** ORTOFRU inizio ***/

        private void FillOrtoFrutticola(string strFil, string strFor, string strDes, string strTip, string strTfo)
        {
            DataRow x;
            string s = "";
            //string sRig = "";

            //s = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + txtFilPF_1.Text + ";Extended Properties=\"Excel 8.0;HDR={1};IMEX=1\"";
            //OleDbConnection cn = new System.Data.OleDb.OleDbConnection(s);

            s = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + strFil + ";Extended Properties=\"Excel 8.0;HDR={1};IMEX=1\"";

            OleDbConnection cn = new System.Data.OleDb.OleDbConnection(s);
            cn.Open();

            DataTable sch = cn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, new object[] { null, null, null, "TABLE" });
            string sTab = (string)sch.Rows[0][2]; 

            OleDbCommand cm = new OleDbCommand();
            cm.Connection = cn;

            //OleDbDataReader dr = cm.ExecuteReader();
            //DataTable sch = dr.GetSchemaTable();
            //s = sch.TableName[0].ToString();
            //s = "SELECT * FROM [Sheet1$]";

            s = "SELECT * FROM [" + sTab + "]";
            cm.CommandText = s;
            OleDbDataReader dr = cm.ExecuteReader();
            DataTable tXls = new DataTable();
            tXls.TableName = "TabXls";
            tXls.Load(dr);

            //using (StreamReader sr = new StreamReader(strFil, System.Text.Encoding.Default))
            //{   

            string sTip = strTip;
            //if (Path.GetFileName(strFil).Substring(0, 1).ToUpper() == "L")
            //    sTip = "LIS";

            clsDivOrtoFrutticola clsDiv = new clsDivOrtoFrutticola();

            //FileInfo fInfo = new FileInfo(strFil);

            progressBar1.Value = 0;
            progressBar1.Maximum = tXls.Rows.Count;
            progressBar1.Minimum = 0;

            DataTable tTes = (DataTable)dgv1.DataSource;

            DataTable tDie = _dasGen.Tables[TABTMPEAN];
            DataColumn[] keys = new DataColumn[2];
            keys[0] = tDie.Columns["die_arf"];
            keys[1] = tDie.Columns["die_ean"];
            tDie.PrimaryKey = keys;

            DataTable t = _dasGen.Tables[TABTMPFDR];

            string sTva = "";
            string sNdo = "";
            DateTime dDdo = _clsDef.DAYOUT;
            string sNum = "";
            string sNums = "";
            string sDes = Path.GetFileNameWithoutExtension(strFil).Replace("\\", "-");
            int iRig = 1;
            DataRow[] j;

            sTva = VARART;
            if (sTip == "OFF")
            {
                //sTva = "OFF";
                dDdo = DateTime.Today;
            }
            else if (sTip == "FAT")
                sTva = "FAT";

            //sNum = (tTes.Rows.Count + 1).ToString("000");
            sNum = DitNewNum();

            foreach(DataRow yy in tXls.Rows)
            {
                progressBar1.Increment(1);
                System.Windows.Forms.Application.DoEvents();

                if (!DBNull.Value.Equals(yy[0]) && _clsFun.Numerico(yy[0]))
                {
                    //sNum = (tTes.Rows.Count).ToString("000");
                    //sNum = DitNewNum();
                    sNum = "001";
                    if (Convert.ToDecimal(tTes.Rows.Count) > 0)
                        sNum = (string)tTes.Rows[tTes.Rows.Count - 1]["dit_num"];

                    sNdo = sNum;
                    //if (sTip == "OFF")
                    //    sNdo = sRig.Substring(22, 5);
                    //else if (sTip == "FAT")
                    //    sNdo = sRig.Substring(1, 7);

                    //s = sRig.Substring(66, 14).Trim();
                    //if (s == "0001134")
                    //    Console.WriteLine("aaaa");
                    //if (s == "0417493")
                    //    Console.WriteLine("aaaa");

                    //j = tTes.Select("dit_tva='" + sTva + "' AND dit_ndo='" + sNdo + "' AND dit_ddo=#" + dDdo.ToString("MM/dd/yyyy") + "#");
                    s = "dit_tva='" + sTva + "' AND dit_des='" + sDes + "'";
                    j = tTes.Select(s);

                    if (j.Length == 0)
                    {
                        //sNum = (tTes.Rows.Count + 1).ToString("000");
                        sNum = DitNewNum();

                        sNums += sNum + ";";

                        x = tTes.NewRow();
                        x["dit_num"] = sNum;
                        x["dit_for"] = strFor;
                        x["dit_des"] = sDes;
                        x["dit_fil"] = Path.GetFileName(strFil);
                        x["dit_tva"] = sTva;

                        x["dit_ndo"] = sNum;    //sNdo
                        if (sTip == "OFF")
                            x["dit_ndo"] = sNdo;
                        else if (sTip == "FAT")
                            x["dit_ndo"] = sNdo;

                        x["dit_ddo"] = DateTime.Today;
                        if (sTva == "FAT")
                            x["dit_ddo"] = DateTime.Today;

                        x["dit_pth"] = Path.GetDirectoryName(strFil);
                        x["dit_nri"] = 1;
                        x["dit_tfo"] = strTfo;
                        x["dit_cho"] = "E";

                        //if (sTip == "OFF")
                        //{
                        //    x["dit_ndo"] = sNdo;

                        //    //s = sRig.Substring(61, 8);
                        //    //s = "20" + s.Substring(6, 2) + s.Substring(3, 2) + s.Substring(0, 2);
                        //    //x["dit_sii"] = _clsFun.Str2Day(s);

                        //    //s = sRig.Substring(6, 8);
                        //    //s = "20" + s.Substring(6, 2) + s.Substring(3, 2) + s.Substring(0, 2);
                        //    //x["dit_sif"] = _clsFun.Str2Day(s);

                        //    //s = sRig.Substring(6, 8);
                        //    //s = "20" + s.Substring(6, 2) + s.Substring(3, 2) + s.Substring(0, 2);
                        //    //x["dit_soi"] = _clsFun.Str2Day(s);
                        //    x["dit_soi"] = DateTime.Today;

                        //    //s = sRig.Substring(14, 8);
                        //    //s = "20" + s.Substring(6, 2) + s.Substring(3, 2) + s.Substring(0, 2);
                        //    //x["dit_sof"] = _clsFun.Str2Day(s);

                        //    //x["dit_des"] = "OFFERTA al " + DateTime.Now.ToShortDateString();

                        //}

                        tTes.Rows.Add(x);
                        iRig = 1;
                        dDdo = (DateTime)x["dit_ddo"];
                    }
                    else
                    {
                        sNum = (string)j[0]["dit_num"];
                    }

                    x = t.NewRow();
                    if (sTip == "VAR")
                        clsDiv.DivLis(yy, x);     /* Lettura record */
                    else if (sTip == "OFF")
                        clsDiv.DivOff(yy, x);     /* Lettura record */
                    //else if (sTip == "FAT")
                    //    clsDiv.DivFat(sRig, x);     /* Lettura record */
                    x["div_tva"] = sTva;
                    x["div_num"] = sNum;
                    x["div_nri"] = iRig.ToString("00000");
                    iRig++;

                    if ((string)x["div_arf"] == "23467")
                        Console.WriteLine("ssss");
                    if ((string)x["div_arf"] == "23887")
                        Console.WriteLine("ssss");

                    j = t.Select("div_tva='" + sTva + "' AND div_num='" + sNum + "' AND div_arf='" + x["div_arf"] + "'");
                    if (j.Length == 0)
                        t.Rows.Add(x);

                    /* Barcode */
                    s = ((string)x["div_ean"]).Trim();
                    if (s != "")
                    {
                        if (s.Length > 7)
                        {
                            if (_strPar020EsclusioneEan2021davDiv == "S" && s.Length == 13 && s.Substring(7, 5) == "00000" && ((s.Substring(0, 2) == "20") || s.Substring(0, 2) == "21"))
                                Console.WriteLine("Scarto i barcode a peso che iniziano per 20 e 21");
                            else
                            {
                                s = s.Substring(0, s.Length - 1) + new clsCtrlCodici().FindMod10Digit(s.Substring(0, s.Length - 1));
                                j = tDie.Select("die_arf='" + x["div_arf"] + "' AND die_ean='" + s + "'");
                                if (j.Length == 0)
                                {
                                    DataRow y = tDie.NewRow();
                                    y["die_num"] = sNum;
                                    y["die_for"] = strFor;
                                    y["die_arf"] = x["div_arf"];
                                    y["die_ean"] = s;
                                    tDie.Rows.Add(y);

                                    if (s != ((string)x["div_ean"]).Trim())
                                        _clsFun.ErrorLog("Divulgazione OrtoFrutticola", "Corretto check digit " + ((string)x["div_ean"]).Trim() + " -> " + s);
                                }
                            }
                        }
                    }
                }
            }

            string[] a = sNums.Split(';');

            for (int i = 0; i < a.Length; i++)
            {
                if (a[i] != "")
                {
                    DataTable tTab = CtrlDivArt(strFil, t, strFor, a[i]);
                    //if (sTip != "OFF")          //Evito controllo articoli
                    //{
                    DataView v = new DataView(tTab, "tab_cod=''", "", DataViewRowState.CurrentRows);
                    if (v.Count > 0)
                    {
                        frmGesDivCtrlTab f = new frmGesDivCtrlTab();
                        f._tabTab = tTab.Copy();
                        f._tabDiv = t.Copy();
                        f._strForCod = strFor;
                        f.ShowDialog();
                        if (f._bolMdy)
                        {
                            CtrlDivArt(strFil, t, strFor, a[i]);
                        }
                    }
                }
                //}
            }
        }

        //private void FillOrtoFrutticolaFat(string strFil, string strFor, string strDes)
        private void FillOrtoFrutticolaFat(string strFil, DataRow rowDiv)
        {
            string strFor = (string)rowDiv["tab_cod"];
            string strDes = (string)rowDiv["tab_des"];
            string strTfo = (string)rowDiv["tab_tip"];
            string strNeg = (string)rowDiv["tab_neg"];

            DataRow x;
            string s = "";
            string sRig = "";

            using (StreamReader sr = new StreamReader(strFil, System.Text.Encoding.Default))
            {
                clsDivOrtoFrutticola clsDiv = new clsDivOrtoFrutticola();

                FileInfo fInfo = new FileInfo(strFil);

                progressBar1.Value = 0;
                progressBar1.Maximum = (int)fInfo.Length;
                progressBar1.Minimum = 0;

                DataTable tTes = (DataTable)dgv1.DataSource;

                DataTable tDie = _dasGen.Tables[TABTMPEAN];
                DataColumn[] keys = new DataColumn[2];
                keys[0] = tDie.Columns["die_arf"];
                keys[1] = tDie.Columns["die_ean"];
                tDie.PrimaryKey = keys;

                DataTable t = _dasGen.Tables[TABTMPFDR];

                string sTva = "FAT";
                string sTri = "";
                string sNdo = "";
                DateTime dDdo = _clsDef.DAYOUT;
                string sNum = "";
                string sNums = "";
                int iRig = 1;
                DataRow[] j;

                while ((sRig = sr.ReadLine()) != null)
                {
                    progressBar1.Increment(sRig.Length);
                    System.Windows.Forms.Application.DoEvents();

                    if (sRig.Length > 2)
                    {
                        sTri = sRig.Substring(4, 1);
                        sNdo = sRig.Substring(5, 8);

                        if (sTri == "F")
                        {
                            dDdo = _clsFun.Str2Day(sRig.Substring(13, 8));
                            j = tTes.Select("dit_tva='" + sTva + "' AND dit_ndo='" + sNdo + "' AND dit_ddo=#" + dDdo.ToString("MM/dd/yyyy") + "#");
                            if (j.Length == 0)
                            {
                                //sNum = (tTes.Rows.Count + 1).ToString("000");
                                sNum = DitNewNum();

                                string sNeg = "001";

                                s = Path.GetFileName(strFil);

                                if (s.Length > 21)
                                {
                                    string[] aNeg = strNeg.Split(',');
                                    if (aNeg.Length > 1)
                                    {
                                        //s = Path.GetFileName(strFil);
                                        //s = s.Substring(15, 6);

                                        string[] aa = s.Split('_');
                                        if (aa.Length > 2)
                                        {
                                            s = aa[2];

                                            for (int ii = 0; ii < aNeg.Length; ii++)
                                            {
                                                if (aNeg[ii] == s)
                                                {
                                                    sNeg = (ii + 1).ToString("000");
                                                    break;
                                                }
                                            }
                                        }
                                    }
                                }
                                if(sNeg == "")
                                    sNeg = "001";




                                sNums += sNum + ";";

                                x = tTes.NewRow();
                                x["dit_num"] = sNum;
                                x["dit_for"] = strFor;
                                x["dit_des"] = strDes;
                                x["dit_fil"] = Path.GetFileName(strFil);
                                x["dit_tva"] = sTva;
                                x["dit_ndo"] = sNdo;
                                x["dit_ddo"] = dDdo;
                                x["dit_pth"] = Path.GetDirectoryName(strFil);
                                x["dit_nri"] = 1;
                                x["dit_cho"] = "E";
                                x["dit_neg"] = sNeg;

                                tTes.Rows.Add(x);
                                iRig = 0;
                                dDdo = (DateTime)x["dit_ddo"];
                            }

                            x = t.NewRow();
                            clsDiv.DivFat(sRig, x);     /***/
                            x["div_tva"] = sTva;
                            x["div_num"] = sNum;
                            x["div_dva"] = dDdo;
                            x["div_nri"] = iRig.ToString("00000");

                            j = t.Select("div_tva='" + sTva + "' AND div_num='" + sNum + "' AND div_arf='" + x["div_arf"] + "'");
                            if (j.Length == 0)
                                t.Rows.Add(x);

                            if ((string)x["div_arf"] == "026200")
                                Console.WriteLine("aaaaaaaaa");
                            s = ((string)x["div_ean"]).Trim();
                            if (s != "" && s.Length > 7 && _clsFun.Numerico(s))
                            {
                                if (_strPar020EsclusioneEan2021davDiv == "S" && s.Length == 13 && s.Substring(7, 5) == "00000" && ((s.Substring(0, 2) == "20") || s.Substring(0, 2) == "21"))
                                    Console.WriteLine("Scarto i barcode a peso che iniziano per 20 e 21");
                                else
                                {
                                    s = s.Substring(0, s.Length - 1) + new clsCtrlCodici().FindMod10Digit(s.Substring(0, s.Length - 1));
                                    j = tDie.Select("die_arf='" + x["div_arf"] + "' AND die_ean='" + s + "'");
                                    if (j.Length == 0)
                                    {
                                        DataRow y = tDie.NewRow();
                                        y["die_num"] = sNum;
                                        y["die_for"] = strFor;
                                        y["die_arf"] = x["div_arf"];
                                        y["die_ean"] = s;
                                        tDie.Rows.Add(y);
                                        if (s != ((string)x["div_ean"]).Trim())
                                            _clsFun.ErrorLog("Divulgazione BRENDOLAN", "Corretto check digit " + ((string)x["div_ean"]).Trim() + " -> " + s);
                                    }
                                }
                            }
                        }
                    }
                    //break;
                    iRig++;
                }

                string[] a = sNums.Split(';');
                for (int i = 0; i < a.Length; i++)
                {

                    DataTable tTab = CtrlDivArt(strFil, t, strFor, a[i]);
                    DataView v = new DataView(tTab, "tab_cod=''", "", DataViewRowState.CurrentRows);
                    if (v.Count > 0)
                    {
                        frmGesDivCtrlTab f = new frmGesDivCtrlTab();
                        f._tabTab = tTab.Copy();
                        f._strForCod = strFor;
                        f.ShowDialog();
                        if (f._bolMdy)
                        {
                            CtrlDivArt(strFil, t, strFor, a[i]);
                        }
                    }
                }
            }
        }

        /*** ORTOFRU fine ***/

        /**** DMO inizio ***/

        private void FillDmo(string strFil, string strFor, string strDes, string strTip, string strTfo)
        {
            DataRow x;
            string s = "";
            //string sRig = "";

            string sIva = _clsFun.ParGet(clsDefine.enuParametri.Par012CodIVAxDefault, _strConSql);

            //s = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + txtFilPF_1.Text + ";Extended Properties=\"Excel 8.0;HDR={1};IMEX=1\"";
            //OleDbConnection cn = new System.Data.OleDb.OleDbConnection(s);

            //s = "Provider=Microsoft.Jet.OLEDB.8.0;Data Source=" + strFil + ";Extended Properties=\"Excel 8.0;HDR={1};IMEX=1\"";
            //@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + filename + ";Extended Properties='Excel 12.0 xml;HDR=YES'";
            s = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + strFil + ";Extended Properties=\"Excel 12.0;HDR={1};IMEX=1\"";

            OleDbConnection cn = new System.Data.OleDb.OleDbConnection(s);
            cn.Open();

            DataTable sch = cn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, new object[] { null, null, null, "TABLE" });
            string sTab = (string)sch.Rows[0][2];

            OleDbCommand cm = new OleDbCommand();
            cm.Connection = cn;

            //OleDbDataReader dr = cm.ExecuteReader();
            //DataTable sch = dr.GetSchemaTable();
            //s = sch.TableName[0].ToString();
            //s = "SELECT * FROM [Sheet1$]";

            s = "SELECT * FROM [" + sTab + "]";
            cm.CommandText = s;
            OleDbDataReader dr = cm.ExecuteReader();
            DataTable tXls = new DataTable();
            tXls.TableName = "TabXls";
            tXls.Load(dr);

            //using (StreamReader sr = new StreamReader(strFil, System.Text.Encoding.Default))
            //{   

            string sTip = strTip;
            //if (Path.GetFileName(strFil).Substring(0, 1).ToUpper() == "L")
            //    sTip = "LIS";

            clsDivDmo clsDiv = new clsDivDmo();

            //FileInfo fInfo = new FileInfo(strFil);

            progressBar1.Value = 0;
            progressBar1.Maximum = tXls.Rows.Count;
            progressBar1.Minimum = 0;

            DataTable tTes = (DataTable)dgv1.DataSource;

            DataTable tDie = _dasGen.Tables[TABTMPEAN];
            DataColumn[] keys = new DataColumn[2];
            keys[0] = tDie.Columns["die_arf"];
            keys[1] = tDie.Columns["die_ean"];
            tDie.PrimaryKey = keys;

            DataTable t = _dasGen.Tables[TABTMPFDR];

            string sTva = "";
            string sNdo = "";
            DateTime dDdo = _clsDef.DAYOUT;
            string sNum = "";
            string sNums = "";
            string sDes = Path.GetFileNameWithoutExtension(strFil).Replace("\\", "-");
            int iRig = 1;
            DataRow[] j;

            sTva = VARART;
            if (sTip == "OFF")
            {
                //sTva = "OFF";
                dDdo = DateTime.Today;
            }
            else if (sTip == "FAT")
                sTva = "FAT";

            //sNum = (tTes.Rows.Count + 1).ToString("000");
            sNum = DitNewNum();

            foreach (DataRow yy in tXls.Rows)
            {
                progressBar1.Increment(1);
                System.Windows.Forms.Application.DoEvents();

                if (!DBNull.Value.Equals(yy[0]) && _clsFun.Numerico(yy[0]))
                {
                    //sNum = (tTes.Rows.Count).ToString("000");
                    //sNum = DitNewNum();
                    sNum = "001";
                    if (Convert.ToDecimal(tTes.Rows.Count) > 0)
                        sNum = (string)tTes.Rows[tTes.Rows.Count - 1]["dit_num"];

                    sNdo = sNum;
                    //if (sTip == "OFF")
                    //    sNdo = sRig.Substring(22, 5);
                    //else if (sTip == "FAT")
                    //    sNdo = sRig.Substring(1, 7);

                    //s = sRig.Substring(66, 14).Trim();
                    //if (s == "0001134")
                    //    Console.WriteLine("aaaa");
                    //if (s == "0417493")
                    //    Console.WriteLine("aaaa");

                    //j = tTes.Select("dit_tva='" + sTva + "' AND dit_ndo='" + sNdo + "' AND dit_ddo=#" + dDdo.ToString("MM/dd/yyyy") + "#");
                    s = "dit_tva='" + sTva + "' AND dit_des='" + sDes + "'";
                    j = tTes.Select(s);

                    if (j.Length == 0)
                    {
                        //sNum = (tTes.Rows.Count + 1).ToString("000");
                        sNum = DitNewNum();

                        sNums += sNum + ";";

                        x = tTes.NewRow();
                        x["dit_num"] = sNum;
                        x["dit_for"] = strFor;
                        x["dit_des"] = sDes;
                        x["dit_fil"] = Path.GetFileName(strFil);
                        x["dit_tva"] = sTva;

                        x["dit_ndo"] = sNum;    //sNdo
                        if (sTip == "OFF")
                            x["dit_ndo"] = sNdo;
                        else if (sTip == "FAT")
                            x["dit_ndo"] = sNdo;

                        x["dit_ddo"] = DateTime.Today;
                        if (sTva == "FAT")
                            x["dit_ddo"] = DateTime.Today;

                        x["dit_pth"] = Path.GetDirectoryName(strFil);
                        x["dit_nri"] = 1;
                        x["dit_tfo"] = strTfo;
                        x["dit_cho"] = "E";

                        //if (sTip == "OFF")
                        //{
                        //    x["dit_ndo"] = sNdo;

                        //    //s = sRig.Substring(61, 8);
                        //    //s = "20" + s.Substring(6, 2) + s.Substring(3, 2) + s.Substring(0, 2);
                        //    //x["dit_sii"] = _clsFun.Str2Day(s);

                        //    //s = sRig.Substring(6, 8);
                        //    //s = "20" + s.Substring(6, 2) + s.Substring(3, 2) + s.Substring(0, 2);
                        //    //x["dit_sif"] = _clsFun.Str2Day(s);

                        //    //s = sRig.Substring(6, 8);
                        //    //s = "20" + s.Substring(6, 2) + s.Substring(3, 2) + s.Substring(0, 2);
                        //    //x["dit_soi"] = _clsFun.Str2Day(s);
                        //    x["dit_soi"] = DateTime.Today;

                        //    //s = sRig.Substring(14, 8);
                        //    //s = "20" + s.Substring(6, 2) + s.Substring(3, 2) + s.Substring(0, 2);
                        //    //x["dit_sof"] = _clsFun.Str2Day(s);

                        //    //x["dit_des"] = "OFFERTA al " + DateTime.Now.ToShortDateString();

                        //}

                        tTes.Rows.Add(x);
                        iRig = 1;
                        dDdo = (DateTime)x["dit_ddo"];
                    }
                    else
                    {
                        sNum = (string)j[0]["dit_num"];
                    }

                    x = t.NewRow();
                    if (sTip == "VAR")
                    {
                        clsDiv.DivLis(yy, x);     /* Lettura record */
                        x["for_iva"] = sIva;
                    }
                    //else if (sTip == "OFF")
                    //    clsDiv.DivOff(yy, x);     /* Lettura record */
                    //else if (sTip == "FAT")
                    //    clsDiv.DivFat(sRig, x);     /* Lettura record */
                    x["div_tva"] = sTva;
                    x["div_num"] = sNum;
                    x["div_nri"] = iRig.ToString("00000");
                    iRig++;

                    if ((string)x["div_arf"] == "23467")
                        Console.WriteLine("ssss");
                    if ((string)x["div_arf"] == "23887")
                        Console.WriteLine("ssss");

                    j = t.Select("div_tva='" + sTva + "' AND div_num='" + sNum + "' AND div_arf='" + x["div_arf"] + "'");
                    if (j.Length == 0)
                        t.Rows.Add(x);

                    /* Barcode */
                    s = ((string)x["div_ean"]).Trim();
                    if (s != "")
                    {
                        if (s.Length > 7)
                        {
                            if (_strPar020EsclusioneEan2021davDiv == "S" && s.Length == 13 && s.Substring(7, 5) == "00000" && ((s.Substring(0, 2) == "20") || s.Substring(0, 2) == "21"))
                                Console.WriteLine("Scarto i barcode a peso che iniziano per 20 e 21");
                            else
                            {
                                s = s.Substring(0, s.Length - 1) + new clsCtrlCodici().FindMod10Digit(s.Substring(0, s.Length - 1));
                                j = tDie.Select("die_arf='" + x["div_arf"] + "' AND die_ean='" + s + "'");
                                if (j.Length == 0)
                                {
                                    DataRow y = tDie.NewRow();
                                    y["die_num"] = sNum;
                                    y["die_for"] = strFor;
                                    y["die_arf"] = x["div_arf"];
                                    y["die_ean"] = s;
                                    tDie.Rows.Add(y);

                                    if (s != ((string)x["div_ean"]).Trim())
                                        _clsFun.ErrorLog("Divulgazione OrtoFrutticola", "Corretto check digit " + ((string)x["div_ean"]).Trim() + " -> " + s);
                                }
                            }
                        }
                    }
                }
            }

            string[] a = sNums.Split(';');

            for (int i = 0; i < a.Length; i++)
            {
                DataTable tTab = CtrlDivArt(strFil, t, strFor, a[i]);
                //if (sTip != "OFF")          //Evito controllo articoli
                //{
                DataView v = new DataView(tTab, "tab_cod=''", "", DataViewRowState.CurrentRows);
                if (v.Count > 0)
                {
                    frmGesDivCtrlTab f = new frmGesDivCtrlTab();
                    f._tabTab = tTab.Copy();
                    f._tabDiv = t.Copy();
                    f._strForCod = strFor;
                    f.ShowDialog();
                    if (f._bolMdy)
                    {
                        CtrlDivArt(strFil, t, strFor, a[i]);
                    }
                }
                //}
            }
        }

        //private void FillDmoFat(string strFil, string strFor, string strDes)
        private void FillDmoFat(string strFil, DataRow rowDiv)
        {
            string strFor = (string)rowDiv["tab_cod"];
            string strDes = (string)rowDiv["tab_des"];
            string strTfo = (string)rowDiv["tab_tip"];
            string strNeg = (string)rowDiv["tab_neg"];

            DataRow x;
            string s = "";

            string sIvaDef = _clsFun.ParGet(clsDefine.enuParametri.Par012CodIVAxDefault, _strConSql);

            using (StreamReader sr = new StreamReader(strFil, System.Text.Encoding.Default))
            {
                clsDivDmo clsDiv = new clsDivDmo();

                FileInfo fInfo = new FileInfo(strFil);

                progressBar1.Value = 0;
                progressBar1.Maximum = (int)fInfo.Length;
                progressBar1.Minimum = 0;

                DataTable tTes = (DataTable)dgv1.DataSource;

                DataTable tDie = _dasGen.Tables[TABTMPEAN];
                DataColumn[] keys = new DataColumn[2];
                keys[0] = tDie.Columns["die_arf"];
                keys[1] = tDie.Columns["die_ean"];
                tDie.PrimaryKey = keys;

                DataTable t = _dasGen.Tables[TABTMPFDR];

                string sTva = "FAT";
                string sTri = "";
                string sNdo = "";
                DateTime dDdo = _clsDef.DAYOUT;
                string sNum = "";
                string sNums = "";
                int iRig = 1;
                int iR = 1;
                DataRow[] j;
                string sRig = "";
                string sR = "";

                //Congiungere 2 righe

                while ((sR = sr.ReadLine()) != null)
                {
                    progressBar1.Increment(sRig.Length);
                    System.Windows.Forms.Application.DoEvents();

                    //iR++;

                    //if (sRig.Length == 0)
                    //{
                    //    sR = sR + new string(' ', 75);
                    //    sR = sR.Substring(0, 75);
                    //}
                    sRig += sR;

                    if (sRig.Length > 100)
                    {
                        sTri = sRig.Substring(81, 1);
                        sNdo = sRig.Substring(83, 17).Trim();

                        if (sTri == "B")
                            sTri = "F";

                        if (sTri == "F")
                        {
                            dDdo = _clsFun.Str2Day(sRig.Substring(101, 8));
                            j = tTes.Select("dit_tva='" + sTva + "' AND dit_ndo='" + sNdo + "' AND dit_ddo=#" + dDdo.ToString("MM/dd/yyyy") + "#");
                            if (j.Length == 0)
                            {
                                string sNeg = "001";

                                string[] aNeg = strNeg.Split(',');
                                if (aNeg.Length > 1)
                                {
                                    s = sRig.Substring(1, 5);
                                    for (int ii = 0; ii < aNeg.Length; ii++)
                                    {
                                        if (aNeg[ii] == s)
                                        {
                                            sNeg = (ii + 1).ToString("000");
                                            break;
                                        }
                                    }
                                }
                                else
                                    sNeg = "001";

                                //sNum = (tTes.Rows.Count + 1).ToString("000");
                                sNum = DitNewNum();

                                sNums += sNum + ";";

                                x = tTes.NewRow();
                                x["dit_num"] = sNum;
                                x["dit_for"] = strFor;
                                x["dit_des"] = strDes;
                                x["dit_fil"] = Path.GetFileName(strFil);
                                x["dit_tva"] = sTva;
                                x["dit_ndo"] = sNdo;
                                x["dit_ddo"] = dDdo;
                                x["dit_pth"] = Path.GetDirectoryName(strFil);
                                x["dit_nri"] = 1;
                                x["dit_neg"] = sNeg;
                                x["dit_cho"] = "E";
                                tTes.Rows.Add(x);
                                iRig = 0;
                                dDdo = (DateTime)x["dit_ddo"];
                            }

                            x = t.NewRow();
                            clsDiv.DivFat(sRig, x);     /***/
                            x["div_tva"] = sTva;
                            x["div_num"] = sNum;
                            x["div_dva"] = dDdo;
                            x["div_nri"] = iRig.ToString("00000");
                            x["div_iva"] = sIvaDef;

                            j = t.Select("div_tva='" + sTva + "' AND div_num='" + sNum + "' AND div_arf='" + x["div_arf"] + "'");
                            if (j.Length == 0)
                                t.Rows.Add(x);

                            if ((string)x["div_arf"] == "026200")
                                Console.WriteLine("aaaaaaaaa");
                            s = ((string)x["div_ean"]).Trim();
                            if (s != "" && s.Length > 7 && _clsFun.Numerico(s))
                            {
                                if (_strPar020EsclusioneEan2021davDiv == "S" && s.Length == 13 && s.Substring(7, 5) == "00000" && ((s.Substring(0, 2) == "20") || s.Substring(0, 2) == "21"))
                                    Console.WriteLine("Scarto i barcode a peso che iniziano per 20 e 21");
                                else
                                {
                                    s = s.Substring(0, s.Length - 1) + new clsCtrlCodici().FindMod10Digit(s.Substring(0, s.Length - 1));
                                    j = tDie.Select("die_arf='" + x["div_arf"] + "' AND die_ean='" + s + "'");
                                    if (j.Length == 0)
                                    {
                                        DataRow y = tDie.NewRow();
                                        y["die_num"] = sNum;
                                        y["die_for"] = strFor;
                                        y["die_arf"] = x["div_arf"];
                                        y["die_ean"] = s;
                                        tDie.Rows.Add(y);
                                        if (s != ((string)x["div_ean"]).Trim())
                                            _clsFun.ErrorLog("Divulgazione BRENDOLAN", "Corretto check digit " + ((string)x["div_ean"]).Trim() + " -> " + s);
                                    }
                                }
                            }
                            iRig++;
                        }
                        sRig = "";
                    }
                    //break;

                }

                string[] a = sNums.Split(';');
                for (int i = 0; i < a.Length; i++)
                {

                    DataTable tTab = CtrlDivArt(strFil, t, strFor, a[i]);
                    DataView v = new DataView(tTab, "tab_cod=''", "", DataViewRowState.CurrentRows);
                    if (v.Count > 0)
                    {
                        frmGesDivCtrlTab f = new frmGesDivCtrlTab();
                        f._tabTab = tTab.Copy();
                        f._strForCod = strFor;
                        f.ShowDialog();
                        if (f._bolMdy)
                        {
                            CtrlDivArt(strFil, t, strFor, a[i]);
                        }
                    }
                }
            }
        }


        /*** DMO fine ***/

        private void FillGabbiano(string strFil, string strFor, string strDes, string strTip, string strTfo)
        {
            DataRow x;
            string s = "";
            //string sRig = "";

            string sIva = _clsFun.ParGet(clsDefine.enuParametri.Par012CodIVAxDefault, _strConSql);
            s = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + strFil + ";Extended Properties=\"Excel 12.0;HDR={1};IMEX=1\"";

            OleDbConnection cn = new System.Data.OleDb.OleDbConnection(s);
            cn.Open();

            DataTable sch = cn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, new object[] { null, null, null, "TABLE" });
            string sTab = (string)sch.Rows[0][2];

            OleDbCommand cm = new OleDbCommand();
            cm.Connection = cn;

            s = "SELECT * FROM [" + sTab + "]";
            cm.CommandText = s;
            OleDbDataReader dr = cm.ExecuteReader();
            DataTable tXls = new DataTable();
            tXls.TableName = "TabXls";
            tXls.Load(dr);

            string sTip = strTip;

            clsDivGabbiano clsDiv = new clsDivGabbiano();

            progressBar1.Value = 0;
            progressBar1.Maximum = tXls.Rows.Count;
            progressBar1.Minimum = 0;

            DataTable tTes = (DataTable)dgv1.DataSource;

            DataTable tDie = _dasGen.Tables[TABTMPEAN];
            DataColumn[] keys = new DataColumn[2];
            keys[0] = tDie.Columns["die_arf"];
            keys[1] = tDie.Columns["die_ean"];
            tDie.PrimaryKey = keys;

            DataTable t = _dasGen.Tables[TABTMPFDR];

            string sTva = "";
            string sNdo = "";
            DateTime dDdo = _clsDef.DAYOUT;
            string sNum = "";
            string sNums = "";
            string sDes = Path.GetFileNameWithoutExtension(strFil).Replace("\\", "-");

            if (sDes.Length > 50)
                sDes = sDes.Substring(0,50);

            int iRig = 1;
            DataRow[] j;

            sTva = VARART;
            if (sTip == "OFF")
            {
                dDdo = DateTime.Today;
            }
            else if (sTip == "FAT")
                sTva = "FAT";

            //sNum = (tTes.Rows.Count + 1).ToString("000");
            sNum = DitNewNum();

            foreach (DataRow yy in tXls.Rows)
            {
                progressBar1.Increment(1);
                System.Windows.Forms.Application.DoEvents();

                s = "";
                if (!DBNull.Value.Equals(yy[3]))
                {
                    try
                    {
                        s = Convert.ToDouble(yy[3]).ToString();
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(ex.Message);
                    }
                }
                if (s.Length > 13)
                    s = "";

                //if (!DBNull.Value.Equals(yy[2]) && _clsFun.Numerico(yy[2]))
                if (s != "")
                {
                    //sNum = (tTes.Rows.Count).ToString("000");
                    sNum = (string)tTes.Rows[tTes.Rows.Count-1]["dit_num"];
                    //sNum = DitNewNum();

                    sNdo = sNum;

                    s = "dit_tva='" + sTva + "' AND dit_des='" + sDes + "'";
                    j = tTes.Select(s);

                    if (j.Length == 0)
                    {
                        //sNum = (tTes.Rows.Count + 1).ToString("000");
                        sNum = DitNewNum();

                        sNums += sNum + ";";

                        x = tTes.NewRow();
                        x["dit_num"] = sNum;
                        x["dit_for"] = strFor;
                        x["dit_des"] = sDes;
                        x["dit_fil"] = Path.GetFileName(strFil);
                        x["dit_tva"] = sTva;

                        x["dit_ndo"] = sNum;    //sNdo
                        if (sTip == "OFF")
                            x["dit_ndo"] = sNdo;
                        else if (sTip == "FAT")
                            x["dit_ndo"] = sNdo;

                        x["dit_ddo"] = DateTime.Today;
                        if (sTva == "FAT")
                            x["dit_ddo"] = DateTime.Today;

                        x["dit_pth"] = Path.GetDirectoryName(strFil);
                        x["dit_nri"] = 1;
                        x["dit_tfo"] = strTfo;
                        x["dit_cho"] = "E";

                        tTes.Rows.Add(x);
                        iRig = 1;
                        dDdo = (DateTime)x["dit_ddo"];
                    }
                    else
                    {
                        sNum = (string)j[0]["dit_num"];
                    }

                    x = t.NewRow();
                    if (sTip == "VAR")
                    {
                        clsDiv.DivLis(yy, x);     /* Lettura record */
                        x["for_iva"] = sIva;
                    }
                    x["div_tva"] = sTva;
                    x["div_num"] = sNum;
                    x["div_nri"] = iRig.ToString("00000");
                    iRig++;

                    if ((string)x["div_arf"] == "23467")
                        Console.WriteLine("ssss");
                    if ((string)x["div_arf"] == "23887")
                        Console.WriteLine("ssss");

                    j = t.Select("div_tva='" + sTva + "' AND div_num='" + sNum + "' AND div_arf='" + x["div_arf"] + "'");
                    if (j.Length == 0)
                        t.Rows.Add(x);

                    /* Barcode */
                    s = ((string)x["div_ean"]).Trim();
                    if (s != "")
                    {
                        if (s.Length > 7)
                        {
                            if (_strPar020EsclusioneEan2021davDiv == "S" && s.Length == 13 && s.Substring(7, 5) == "00000" && ((s.Substring(0, 2) == "20") || s.Substring(0, 2) == "21"))
                                Console.WriteLine("Scarto i barcode a peso che iniziano per 20 e 21");
                            else
                            {
                                s = s.Substring(0, s.Length - 1) + new clsCtrlCodici().FindMod10Digit(s.Substring(0, s.Length - 1));
                                j = tDie.Select("die_arf='" + x["div_arf"] + "' AND die_ean='" + s + "'");
                                if (j.Length == 0)
                                {
                                    DataRow y = tDie.NewRow();
                                    y["die_num"] = sNum;
                                    y["die_for"] = strFor;
                                    y["die_arf"] = x["div_arf"];
                                    y["die_ean"] = s;
                                    tDie.Rows.Add(y);

                                    if (s != ((string)x["div_ean"]).Trim())
                                        _clsFun.ErrorLog("Divulgazione OrtoFrutticola", "Corretto check digit " + ((string)x["div_ean"]).Trim() + " -> " + s);
                                }
                            }
                        }
                    }
                }
            }

            string[] a = sNums.Split(';');

            for (int i = 0; i < a.Length; i++)
            {
                DataTable tTab = CtrlDivArt(strFil, t, strFor, a[i]);
                DataView v = new DataView(tTab, "tab_cod=''", "", DataViewRowState.CurrentRows);
                if (v.Count > 0)
                {
                    frmGesDivCtrlTab f = new frmGesDivCtrlTab();
                    f._tabTab = tTab.Copy();
                    f._tabDiv = t.Copy();
                    f._strForCod = strFor;
                    f.ShowDialog();
                    if (f._bolMdy)
                    {
                        CtrlDivArt(strFil, t, strFor, a[i]);
                    }
                }
            }
        }

        private void FillUnicom(string strFil, string strTip, DataRow rowDiv)
        {
            string strFor = (string)rowDiv["tab_cod"];
            string strDes = (string)rowDiv["tab_des"];
            string strTfo = (string)rowDiv["tab_tip"];
            string strNeg = (string)rowDiv["tab_neg"];

            DataRow x;
            string s = "";
            string sRig = "";
            string sRig1 = "";
            string sNeg = "";

            using (StreamReader sr = new StreamReader(strFil, System.Text.Encoding.Default))
            {
                clsDivUnicom clsDiv = new clsDivUnicom();

                FileInfo fInfo = new FileInfo(strFil);

                progressBar1.Value = 0;
                progressBar1.Minimum = 0;
                progressBar1.Maximum = (int)fInfo.Length;

                DataTable tTes = (DataTable)dgv1.DataSource;

                DataTable tDie = _dasGen.Tables[TABTMPEAN];
                DataColumn[] keys = new DataColumn[2];
                keys[0] = tDie.Columns["die_arf"];
                keys[1] = tDie.Columns["die_ean"];
                tDie.PrimaryKey = keys;

                DataTable t = _dasGen.Tables[TABTMPFDR];

                string sTva = "";
                string sNdo = "";
                DateTime dDdo = DateTime.Today;
                string sNum = "";
                string sNums = "";
                int iRig = 1;
                string sDes = "";
                DataRow[] j;
                Boolean b = true;

                //sNum = (tTes.Rows.Count + 1).ToString("000");
                sNum = DitNewNum();

                while ((sRig = sr.ReadLine()) != null)
                {
                    progressBar1.Increment(sRig.Length);
                    System.Windows.Forms.Application.DoEvents();

                    b = true;

                    //if(strTip == "FAT")
                    //{
                    //if (sRig.Length > 10 && sRig.Substring(0, 1) == "9")
                    //{
                    //    b = false;
                    //    sRig1 = sRig;
                    //}
                    //else
                    //{
                    //    sRig = sRig1 + " " + sRig;
                    //    sRig1 = "";
                    //}
                    //}

                    if (sRig.Length < 100)
                        b = false;
                    if (sRig.Length > 246)
                        b = false;

                    if (b && sRig.Substring(33, 10).Trim() != "" && sRig.Substring(48, 50).Trim() == "")
                        b = false;

                    if (b && strTip == "VAR" && sRig.Length >= 245 && sRig.Substring(245, 1) == "X")    //Componenti
                        b = false;
                    if (b && strTip == "VAR" && sRig.Length >= 245 && sRig.Substring(245, 1) == "F")    //Fidelity
                        b = false;

                    if (b)
                    {
                        if (strTip == "VAR" && sRig.Length < 246)
                            b = false;
                        else if (b && strTip == "VAR" && sRig.Length > 10 && sRig.Substring(245, 1) == "M")
                        {
                            s = sRig.Substring(11, 3);

                            //if (_clsFun.Numerico(s) && Convert.ToInt16(s) > 10)    //Controllo sul settore merceologico
                            //    b = false;

                            s = sRig.Substring(36, 40).Trim();
                            if (s.Length < 2)
                                b = false;
                        }

                        //sNum = (tTes.Rows.Count).ToString("000");
                        //sNum = DitNewNum();
                        sNum = "001";
                        if (Convert.ToDecimal(tTes.Rows.Count) > 0)
                            sNum = (string)tTes.Rows[tTes.Rows.Count - 1]["dit_num"];

                        if ((strTip == "FAT" || strTip == "BOL") && "BFA".Contains(sRig.Substring(20, 1)) && sRig.Length > 10 && sRig.Substring(0, 1) == "9")
                        {
                            if (sNdo != sRig.Substring(1, 10))
                            {
                                sTva = VARFAT;
                                if (sRig.Substring(20, 1) == "A")
                                    sTva = "NAC";

                                sNeg = "001";

                                string[] aNeg = strNeg.Split(',');
                                if (aNeg.Length > 1)
                                {
                                    s = sRig.Substring(27, 6);
                                    for (int ii = 0; ii < aNeg.Length; ii++)
                                    {
                                        if (aNeg[ii] == s)
                                        {
                                            sNeg = (ii + 1).ToString("000");
                                            break;
                                        }
                                    }
                                }
                                else
                                    sNeg = "001";

                                s = Path.GetFileName(strFil);

                                if (sRig.Substring(20, 1) == "B")
                                    sDes = "Bolla n. " + sRig.Substring(1, 10);
                                else if (sRig.Substring(20, 1) == "A")
                                    sDes = "N.Accredito n. " + sRig.Substring(1, 10);
                                else
                                    sDes = "Fattura n. " + sRig.Substring(1, 10);

                                sNdo = sRig.Substring(1, 10);
                                dDdo = _clsFun.Str2Day(sRig.Substring(11, 8));
                                //sRig1 = sRig;
                            }
                        }
                    }
                    if (b)
                    {
                        if (sRig.Length > 10)
                        {
                            string sTip = "";
                            Boolean bTes = false;

                            if (strTip == "BOL")
                            {
                                if (sRig.Substring(20, 1) == "B")
                                    bTes = true;
                                sTip = "BOL";
                            }
                            else if (strTip == "FAT")
                            {
                                if (sRig.Substring(20, 1) == "F")
                                    bTes = true;
                                sTip = "FAT";

                                if (sRig.Substring(20, 1) == "A")
                                {
                                    bTes = true;
                                    sTip = "NAC";
                                }
                            }
                            else if (sRig.Substring(245, 1) == "T")
                            {
                                bTes = true;
                                sTip = "T";
                            }
                            if (bTes)
                            {

                                if (sTip != "BOL" && sTip != "FAT" && sTip != "NAC")
                                {
                                    sDes = sRig.Substring(36, 40).Trim();
                                    sTva = VARART;

                                    if (sDes.Substring(0, 7).ToLower() == "offerta")
                                        sTva = "OFF";
                                    else
                                    {
                                        dDdo = DateTime.Today;

                                        s = sDes.Substring(sDes.Length - 8);

                                        if (s.Substring(2, 1) == "/")
                                        {
                                            s = "20" + s.Substring(6, 2) + s.Substring(3, 2) + s.Substring(0, 2);
                                            dDdo = _clsFun.Str2Day(s);
                                        }
                                    }

                                    sNdo = sNum;
                                }

                                j = tTes.Select("dit_tva='" + sTva + "' AND dit_ndo='" + sNdo + "' AND dit_ddo=#" + dDdo.ToString("MM/dd/yyyy") + "#");

                                if (j.Length == 0)
                                {
                                    //sNum = (tTes.Rows.Count + 1).ToString("000");
                                    sNum = DitNewNum();
                                    sNums += sNum + ";";

                                    x = tTes.NewRow();
                                    x["dit_num"] = sNum;
                                    x["dit_for"] = strFor;

                                    s = (strDes + " " + sDes + new string(' ', 50)).Substring(0, 50).Trim();
                                    x["dit_des"] = s;

                                    x["dit_fil"] = Path.GetFileName(strFil);
                                    x["dit_tva"] = sTva;
                                    x["dit_ndo"] = sNdo;
                                    x["dit_ddo"] = dDdo;
                                    x["dit_pth"] = Path.GetDirectoryName(strFil);
                                    x["dit_nri"] = 1;
                                    x["dit_tfo"] = strTfo;
                                    x["dit_neg"] = sNeg;
                                    x["dit_cho"] = "E";

                                    tTes.Rows.Add(x);
                                    iRig = 1;
                                }
                            }
                            else
                            {
                                if (strTip != "BOL" && sTip != "FAT" && sTip != "NAC")
                                    sTip = sRig.Substring(245, 1);

                                if (sTip == "O")
                                {
                                    j = tTes.Select("dit_num='" + sNum + "'");
                                    if (j.Length > 0)
                                    {
                                        if (DBNull.Value.Equals(j[0]["dit_soi"]))
                                        {
                                            //s = sRig.Substring(61, 8);
                                            //s = "20" + s.Substring(6, 2) + s.Substring(3, 2) + s.Substring(0, 2);
                                            //x["dit_sii"] = _clsFun.Str2Day(s);

                                            //s = sRig.Substring(6, 8);
                                            //s = "20" + s.Substring(6, 2) + s.Substring(3, 2) + s.Substring(0, 2);
                                            //x["dit_sif"] = _clsFun.Str2Day(s);

                                            s = sRig.Substring(184, 8);
                                            j[0]["dit_soi"] = _clsFun.Str2Day(s);
                                            j[0]["dit_sii"] = j[0]["dit_soi"];

                                            s = sRig.Substring(192, 8);
                                            j[0]["dit_sof"] = _clsFun.Str2Day(s);
                                            j[0]["dit_sif"] = j[0]["dit_sof"];
                                        }
                                    }
                                }
                                else if (strTip != "BOL" && sTip != "FAT" && sTip != "NAC" && "X".Contains(sRig.Substring(245, 1)))
                                    Console.WriteLine("Salta");
                                if (strTip != "BOL" && sTip != "FAT" && sTip != "NAC" && sRig.Substring(245, 1) == "L")
                                {
                                    s = sRig.Substring(4, 7).Trim();

                                    j = t.Select("div_tva='" + sTva + "' AND div_num='" + sNum + "' AND div_arf='" + s + "'");
                                    if (j.Length > 0)
                                        clsDiv.DivIng(sRig, j[0]);
                                }
                                else if (strTip != "BOL" && sTip != "FAT" && sTip != "NAC" && sRig.Substring(245, 1) == "D")
                                {
                                    s = sRig.Substring(4, 7).Trim();

                                    j = t.Select("div_tva='" + sTva + "' AND div_num='" + sNum + "' AND div_arf='" + s + "'");
                                    if (j.Length > 0)
                                        clsDiv.DivBil(sRig, j[0]);
                                }
                                else
                                {
                                    if (iRig == 838)
                                        Console.WriteLine("aaaaaaaa");

                                    x = t.NewRow();
                                    if (sTip == "BOL")
                                        clsDiv.DivBol(sRig, x);     /* Lettura record */
                                    else if (sTip == "FAT" || sTip == "NAC")
                                    {
                                        //s = sRig1 + "|" + sRig;
                                        clsDiv.DivFat(sRig, x);     /* Lettura record */
                                    }
                                    else if ("MAO".Contains(sRig.Substring(245, 1)))
                                        clsDiv.DivArt(sRig, x, dDdo);     /* Lettura record */

                                    if ((string)x["div_arf"] == "2145761")
                                        Console.WriteLine("aaaaa");

                                    x["div_tva"] = sTva;
                                    x["div_num"] = sNum;
                                    x["div_nri"] = iRig.ToString("00000");
                                    iRig++;

                                    //if ((string)x["div_arf"] == "2010010")
                                    //    Console.WriteLine("aaaaaaaaaaaa");

                                    if (sTva == "FAT" || sTva == "NAC" || sTva == "BOL")
                                    {
                                        j = t.Select("div_tva='" + sTva + "' AND div_num='" + sNum + "' AND div_arf='" + x["div_arf"] + "' AND div_nri='" + x["div_nri"] + "'");
                                        if(j.Length == 0)
                                            t.Rows.Add(x);
                                        Console.WriteLine("aaa");
                                    }
                                    else
                                    {
                                        j = t.Select("div_tva='" + sTva + "' AND div_num='" + sNum + "' AND div_arf='" + x["div_arf"] + "'");
                                        if (j.Length == 0)
                                            t.Rows.Add(x);
                                        else
                                        {
                                            if (DBNull.Value.Equals(x["div_plu"]))
                                                x["div_plu"] = "";
                                            if (DBNull.Value.Equals(x["div_ing"]))
                                                x["div_ing"] = "";
                                            if (DBNull.Value.Equals(x["div_bil"]))
                                                x["div_bil"] = "";
                                            if (DBNull.Value.Equals(x["div_reb"]))
                                                x["div_reb"] = "";

                                            if (((string)x["div_plu"]).Trim() != "")
                                                j[0]["div_plu"] = x["div_plu"];
                                            if (((string)x["div_ing"]).Trim() != "")
                                                j[0]["div_ing"] = x["div_ing"];
                                            if (((string)x["div_bil"]).Trim() != "")
                                                j[0]["div_bil"] = x["div_bil"];
                                            if (((string)x["div_reb"]).Trim() != "")
                                                j[0]["div_reb"] = x["div_reb"];
                                        }
                                    }
                                    /* Barcode */
                                    s = ((string)x["div_ean"]).Trim();

                                    if (_clsFun.Numerico(s))
                                    {
                                        s = Convert.ToDouble(s).ToString();

                                        if (s != "")
                                        {
                                            if (s.Length > 7)
                                            {
                                                if (s.Length == 13 && s.Substring(0, 1) == "2" && s.Substring(7, 5) == "00000")
                                                    s = s.Substring(0, 12) + "0";
                                                else
                                                    s = s.Substring(0, s.Length - 1) + new clsCtrlCodici().FindMod10Digit(s.Substring(0, s.Length - 1));

                                                if (_strPar020EsclusioneEan2021davDiv == "S" && s.Length == 13 && s.Substring(7, 5) == "00000" && ((s.Substring(0, 2) == "20") || s.Substring(0, 2) == "21"))
                                                    Console.WriteLine("Scarto i barcode a peso che iniziano per 20 e 21");
                                                else
                                                {
                                                    j = tDie.Select("die_arf='" + x["div_arf"] + "' AND die_ean='" + s + "'");
                                                    if (j.Length == 0)
                                                    {
                                                        DataRow y = tDie.NewRow();
                                                        y["die_num"] = sNum;
                                                        y["die_for"] = strFor;
                                                        y["die_arf"] = x["div_arf"];
                                                        y["die_ean"] = s;
                                                        tDie.Rows.Add(y);

                                                        if (s != ((string)x["div_ean"]).Trim())
                                                            _clsFun.ErrorLog("Divulgazione DPIU", "Corretto check digit " + ((string)x["div_ean"]).Trim() + " -> " + s);
                                                    }
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                }

                string[] a = sNums.Split(';');

                //da controllare
                for (int i = 0; i < a.Length; i++)
                {
                    if (a[i].Trim() != "")
                    {
                        DataTable tTab = CtrlDivArt(strFil, t, strFor, a[i]);
                        //if (sTip != "OFF")          //Evito controllo articoli
                        //{
                        DataView v = new DataView(tTab, "tab_cod=''", "", DataViewRowState.CurrentRows);
                        if (v.Count > 0)
                        {
                            frmGesDivCtrlTab f = new frmGesDivCtrlTab();
                            f._tabTab = tTab.Copy();
                            f._tabDiv = t.Copy();
                            f._strForCod = strFor;
                            f.ShowDialog();
                            if (f._bolMdy)
                            {
                                CtrlDivArt(strFil, t, strFor, a[i]);
                            }
                        }
                    }
                }
            }
        }

        private void FillUnicomPescheria(String strFor, String strFil)
        {
            string s = "";
            DataRow x;
            DataRow[] j;
            string p = "AnaIttico";
            string sRig = "";
            Boolean bFirst = true;

            ArrayList aWhe = new ArrayList();
            aWhe.Add("itt_art");
            ArrayList aExl = new ArrayList();

            s = "SELECT * FROM AnaIttico ";
            DataTable tItt = _clsFun.FillTabSql(p, s, false, _strConSql);
            DataColumn[] keys = new DataColumn[1];
            keys[0] = tItt.Columns["itt_art"];
            tItt.PrimaryKey = keys;

            //s = "SELECT * FROM GesLisAcquisto WHERE acq_for='" + strFor + "' ";

            s = "SELECT ";
            s += "lia_art, lia_arf ";
            s += "FROM GesLisAcquisto ";
            s += "WHERE (lia_for = '" + strFor + "') AND lia_ann=0 ";
            s += "GROUP BY lia_art, lia_arf ";
            s += "ORDER BY lia_arf";
            DataTable tLia = _clsFun.FillTabSql(p, s, false, _strConSql);
            keys = new DataColumn[2];
            keys[0] = tLia.Columns["lia_arf"];
            keys[1] = tLia.Columns["lia_art"];
            tLia.PrimaryKey = keys;

            using (StreamReader sr = new StreamReader(strFil, System.Text.Encoding.Default))
            {
                clsDivUnicom clsDiv = new clsDivUnicom();

                string sFld = "";

                x = tItt.NewRow();

                while ((sRig = sr.ReadLine()) != null)
                {
                    sFld = "itt_0" + sRig.Substring(58, 2);
                    x[sFld] = sRig.Substring(60).Trim();
                    s = sRig;
                }

                x["itt_arf"] = s.Substring(3, 7).Trim();

                j = tLia.Select("lia_arf='" + x["itt_arf"] + "'");
                if (j.Length > 0)
                {
                    string sArt = (string)j[0]["lia_art"];

                    x["itt_art"] = sArt;

                    j = tItt.Select("itt_art='" + sArt + "'");
                    if (j.Length > 0)
                        s = _clsFun.SqlUpdRow(p, tItt, j[0], x, aWhe, aExl);
                    else
                        s = _clsFun.SqlInsertRow(p, tItt, x);
                    if(s != "")
                        _clsFun.SqlWrite(s, _strConSql);

                    _clsVar.Variazioni(sArt, "Forza", "Divulg. Pescheria", _clsDef.VARETI);

                }

                sr.Close();
                sr.Dispose();
            }

            s = Path.GetDirectoryName(strFil) + "\\Old\\" + Path.GetFileName(strFil) + "_" + DateTime.Now.ToString("yyyyMMddHHmm");
            if (File.Exists(s))
                File.Delete(s);
            File.Move(strFil, s);

            //itt_art	nchar(7)	Unchecked
            //itt_arf	nchar(7)	Unchecked
            //itt_001	nvarchar(100)	Unchecked
            //itt_002	nvarchar(100)	Unchecked
            //itt_003	nvarchar(100)	Unchecked
            //itt_004	nvarchar(100)	Unchecked
            //itt_005	nvarchar(100)	Unchecked
            //itt_006	nvarchar(100)	Unchecked
            //itt_007	nvarchar(100)	Unchecked
            //itt_008	nvarchar(100)	Unchecked
            //itt_009	nvarchar(100)	Unchecked
            //itt_010	nvarchar(100)	Unchecked
            //itt_011	nvarchar(100)	Unchecked
            //itt_012	nvarchar(100)	Unchecked

        }

        private void FillGottardo(string strFil, string strTip, DataRow rowDiv)
        {
            DataRow x;
            string s = "";
            string sRig = "";

            //(string)y["tab_cod"], ((string)y["tab_des"]).Trim(), "VAR", (string)y["tab_tip"]);

            string strFor = (string)rowDiv["tab_cod"];
            string strDes = (string)rowDiv["tab_des"];
            string strTfo = (string)rowDiv["tab_tip"];
            string strNeg = (string)rowDiv["tab_neg"];

            using (StreamReader sr = new StreamReader(strFil, System.Text.Encoding.Default))
            {
                clsDivGottardo clsDiv = new clsDivGottardo();
                clsDiv._strPar017EcrDivFornitori = _strPar017EcrDivFornitori;

                FileInfo fInfo = new FileInfo(strFil);

                progressBar1.Value = 0;
                progressBar1.Minimum = 0;
                progressBar1.Maximum = (int)fInfo.Length;

                DataTable tTes = (DataTable)dgv1.DataSource;

                DataTable tDie = _dasGen.Tables[TABTMPEAN];
                DataColumn[] keys = new DataColumn[2];
                keys[0] = tDie.Columns["die_arf"];
                keys[1] = tDie.Columns["die_ean"];
                tDie.PrimaryKey = keys;

                DataTable t = _dasGen.Tables[TABTMPFDR];

                string sTva = "";
                string sNdo = "";
                DateTime dDdo = DateTime.Today;
                string sNum = "";
                string sNums = "";
                int iRig = 1;
                string sDes = "";
                DataRow[] j;

                Boolean bTes = true;

                sNum = DitNewNum();

                sTva = VARART;

                while ((sRig = sr.ReadLine()) != null)
                {
                    progressBar1.Increment(sRig.Length);
                    System.Windows.Forms.Application.DoEvents();

                    if (sRig.Length > 50)
                    {
                        sNum = "001";
                        if (Convert.ToDecimal(tTes.Rows.Count) > 0)
                            sNum = (string)tTes.Rows[tTes.Rows.Count - 1]["dit_num"];

                        if (strTip == "FAT")
                        {
                            s = sRig.Substring(20, 10);

                            if (sNdo != s)
                            {
                                bTes = true;
                                sTva = VARFAT;
                                sDes = "Fattura n. " + s;
                                sNdo = s;
                                dDdo = _clsFun.Str2Day(sRig.Substring(31, 8));
                            }
                        }

                        if (bTes)
                        {
                            string sNeg = "";
                            sNum = DitNewNum();
                            sNums += sNum + ";";

                            if (strTip != "FAT")
                                sDes = "Variazioni Gottardo";
                            else
                            {
                                sNeg = "001";

                                string[] aNeg = strNeg.Split(',');
                                if (aNeg.Length > 1)
                                {
                                    s = Path.GetFileName(strFil);

                                    if (s.Length > 10)
                                    {
                                        s = s.Substring(3, 5);

                                        for (int ii = 0; ii < aNeg.Length; ii++)
                                        {
                                            if (aNeg[ii] == s)
                                            {
                                                sNeg = (ii + 1).ToString("000");
                                                break;
                                            }
                                        }
                                    }
                                }
                            }

                            x = tTes.NewRow();
                            x["dit_num"] = sNum;
                            x["dit_neg"] = sNeg;
                            x["dit_for"] = strFor;
                            x["dit_des"] = sDes;
                            x["dit_fil"] = Path.GetFileName(strFil);
                            x["dit_tva"] = sTva;
                            x["dit_ndo"] = sNdo;
                            x["dit_ddo"] = dDdo;
                            x["dit_pth"] = Path.GetDirectoryName(strFil);
                            x["dit_nri"] = 1;
                            x["dit_tfo"] = strTfo;
                            x["dit_cho"] = "E";

                            tTes.Rows.Add(x);
                            iRig = 1;

                            bTes = false;
                        }

                        x = t.NewRow();
                        if (strTip == "FAT")
                            clsDiv.DivFat(sRig, x);     /* Lettura record */
                        else
                            clsDiv.DivArt(sRig, x);     /* Lettura record */

                        x["div_tva"] = sTva;
                        x["div_num"] = sNum;
                        x["div_nri"] = iRig.ToString("00000");
                        iRig++;

                        j = t.Select("div_tva='" + sTva + "' AND div_num='" + sNum + "' AND div_arf='" + x["div_arf"] + "'");
                        if (j.Length == 0)
                            t.Rows.Add(x);

                        /* Barcode */
                        ArrayList aEan = new ArrayList();

                        if (strTip == "FAT")
                        {
                            s = sRig.Substring(139);

                            int i = 0;
                            while (true)
                            {
                                if (i + 13 > s.Length)
                                    break;

                                string sEan = s.Substring(i, 13);

                                if (sEan == new string('0', 13))
                                    break;
                                else
                                {
                                    sEan = Convert.ToDouble(sEan).ToString();
                                    aEan.Add(sEan);
                                    i += 13;
                                }
                            }
                        }
                        else
                            aEan.Add((string)x["div_ean"]);

                        //s = ((string)x["div_ean"]).Trim();

                        foreach (string sEan in aEan)
                        {
                            if (sEan != "")
                            {
                                if (sEan.Length > 7)
                                {
                                    if (_strPar020EsclusioneEan2021davDiv == "S" && s.Length == 13 && s.Substring(7, 5) == "00000" && ((s.Substring(0, 2) == "20") || s.Substring(0, 2) == "21"))
                                        Console.WriteLine("Scarto i barcode a peso che iniziano per 20 e 21");
                                    else
                                    {
                                        s = sEan.Substring(0, sEan.Length - 1) + new clsCtrlCodici().FindMod10Digit(sEan.Substring(0, sEan.Length - 1));
                                        j = tDie.Select("die_arf='" + x["div_arf"] + "' AND die_ean='" + s + "'");
                                        if (j.Length == 0)
                                        {
                                            DataRow y = tDie.NewRow();
                                            y["die_num"] = sNum;
                                            y["die_for"] = strFor;
                                            y["die_arf"] = x["div_arf"];
                                            y["die_ean"] = s;
                                            tDie.Rows.Add(y);

                                            if (s != sEan)
                                                _clsFun.ErrorLog("Divulgazione Gottardo", "Corretto check digit " + sEan + " -> " + s);
                                        }
                                    }
                                }
                            }
                        }
                    }
                }

                string[] a = sNums.Split(';');

                //da controllare
                for (int i = 0; i < a.Length; i++)
                {
                    if (a[i].Trim() != "")
                    {
                        DataTable tTab = CtrlDivArt(strFil, t, strFor, a[i]);
                        if (strTip != "OFF")          //Evito controllo articoli
                        {
                            DataView v = new DataView(tTab, "tab_cod=''", "", DataViewRowState.CurrentRows);
                            if (v.Count > 0)
                            {
                                frmGesDivCtrlTab f = new frmGesDivCtrlTab();
                                f._tabTab = tTab.Copy();
                                f._tabDiv = t.Copy();
                                f._tabDie = tDie.Copy();
                                f._strForCod = strFor;
                                f.ShowDialog();
                                if (f._bolMdy)
                                {
                                    CtrlDivArt(strFil, t, strFor, a[i]);
                                }
                            }
                        }
                    }
                }
            }
        }

        private void FillFietta(string strFil, string strTip, DataRow rowImp)
        {
            //FillFietta(sFil, (string)y["tab_cod"], ((string)y["tab_des"]).Trim(), "FAT", (string)y["tab_tip"]);

            string strFor = (string)rowImp["tab_cod"];
            string strDes = (string)rowImp["tab_des"];
            string strTfo = (string)rowImp["tab_tip"];
            string strNeg = (string)rowImp["tab_neg"];

            DataRow x;
            string s = "";
            string sRig = "";

            using (StreamReader sr = new StreamReader(strFil, System.Text.Encoding.Default))
            {
                clsDivFietta clsDiv = new clsDivFietta();

                FileInfo fInfo = new FileInfo(strFil);

                progressBar1.Value = 0;
                progressBar1.Minimum = 0;
                progressBar1.Maximum = (int)fInfo.Length;

                DataTable tTes = (DataTable)dgv1.DataSource;

                DataTable tDie = _dasGen.Tables[TABTMPEAN];
                DataColumn[] keys = new DataColumn[2];
                keys[0] = tDie.Columns["die_arf"];
                keys[1] = tDie.Columns["die_ean"];
                tDie.PrimaryKey = keys;

                DataTable t = _dasGen.Tables[TABTMPFDR];

                string sTva = "";
                string sNdo = "";
                DateTime dDdo = DateTime.Today;
                //DateTime dDdo = _clsDef.DAYOUT;
                string sNum = "";
                string sNums = "";
                int iRig = 1;
                string sDes = "";
                DataRow[] j;
                string sNeg = "001";

                Boolean bTes = true;

                //sNum = (tTes.Rows.Count + 1).ToString("000");
                sNum = DitNewNum();

                sTva = VARART;

                while ((sRig = sr.ReadLine()) != null)
                {
                    progressBar1.Increment(sRig.Length);
                    System.Windows.Forms.Application.DoEvents();

                    //sNum = (tTes.Rows.Count).ToString("000");
                    //sNum = DitNewNum();
                    sNum = "001";
                    if (Convert.ToDecimal(tTes.Rows.Count) > 0)
                        sNum = (string)tTes.Rows[tTes.Rows.Count - 1]["dit_num"];

                    if (strTip == "FAT" && sRig.Substring(0, 1) == "T")
                    {
                        bTes = true;
                        sTva = VARFAT;
                        sNdo = sRig.Substring(1, 7);
                        sDes = "Fattura n. " + sNdo;
                        dDdo = _clsFun.Str2Day(sRig.Substring(8, 8));

                        string[] aNeg = strNeg.Split(',');
                        if (aNeg.Length > 1)
                        {
                            s = sRig.Substring(64, 6);
                            for (int ii = 0; ii < aNeg.Length; ii++)
                            {
                                if (aNeg[ii] == s)
                                {
                                    sNeg = (ii + 1).ToString("000");
                                    break;
                                }
                            }
                        }
                    }
                    else if (strTip == "FAT" && sRig.Substring(0, 2) == "01")
                    {
                        bTes = true;
                        sTva = VARFAT;
                        sNdo = sRig.Substring(19, 6);
                        sDes = "Fattura n. " + sNdo;
                        dDdo = _clsFun.Str2Day(sRig.Substring(25, 6));

                        string[] aNeg = strNeg.Split(',');
                        if (aNeg.Length > 1)
                        {
                            s = sRig.Substring(86, 6);
                            for (int ii = 0; ii < aNeg.Length; ii++)
                            {
                                if (aNeg[ii] == s)
                                {
                                    sNeg = (ii + 1).ToString("000");
                                    break;
                                }
                            }
                        }
                    }

                    if (bTes)
                    {
                        sNum = DitNewNum();

                        sNums += sNum + ";";

                        if (strTip != "FAT")
                            sDes = "Variazioni";

                        x = tTes.NewRow();
                        x["dit_num"] = sNum;
                        x["dit_for"] = strFor;
                        x["dit_des"] = sDes;
                        x["dit_fil"] = Path.GetFileName(strFil);
                        x["dit_tva"] = sTva;
                        x["dit_ndo"] = sNdo;
                        x["dit_ddo"] = dDdo;
                        x["dit_pth"] = Path.GetDirectoryName(strFil);
                        x["dit_nri"] = 1;
                        x["dit_tfo"] = strTfo;
                        x["dit_neg"] = sNeg;
                        x["dit_cho"] = "E";

                        tTes.Rows.Add(x);
                        iRig = 1;

                        bTes = false;
                    }

                    else if (sRig.Substring(0, 2) == "02")
                    {
                        x = t.NewRow();
                        if (strTip == "FAT")
                            clsDiv.DivFat(sRig, x, dDdo);     /* Lettura record */
                        else
                            clsDiv.DivArt(sRig, x);     /* Lettura record */

                        x["div_tva"] = sTva;
                        x["div_num"] = sNum;
                        x["div_nri"] = iRig.ToString("00000");
                        iRig++;

                        j = t.Select("div_tva='" + sTva + "' AND div_num='" + sNum + "' AND div_arf='" + x["div_arf"] + "'");
                        if (j.Length == 0)
                            t.Rows.Add(x);

                        string sEan = (string)x["div_ean"];

                        if (sEan != "")
                        {
                            if (sEan.Length > 7)
                            {
                                s = sEan;
                                if (_strPar020EsclusioneEan2021davDiv == "S" && s.Length == 13 && s.Substring(7, 5) == "00000" && ((s.Substring(0, 2) == "20") || s.Substring(0, 2) == "21"))
                                    Console.WriteLine("Scarto i barcode a peso che iniziano per 20 e 21");
                                else
                                {
                                    s = sEan.Substring(0, sEan.Length - 1) + new clsCtrlCodici().FindMod10Digit(sEan.Substring(0, sEan.Length - 1));
                                    j = tDie.Select("die_arf='" + x["div_arf"] + "' AND die_ean='" + s + "'");
                                    if (j.Length == 0)
                                    {
                                        DataRow y = tDie.NewRow();
                                        y["die_num"] = sNum;
                                        y["die_for"] = strFor;
                                        y["die_arf"] = x["div_arf"];
                                        y["die_ean"] = s;
                                        tDie.Rows.Add(y);

                                        if (s != sEan)
                                            _clsFun.ErrorLog("Divulgazione DPIU", "Corretto check digit " + ((string)x["div_ean"]).Trim() + " -> " + s);
                                    }
                                }
                            }
                        }
                    }

                    else if (sRig.Substring(0, 1) == "D")
                    {
                        x = t.NewRow();
                        if (strTip == "FAT")
                        {
                            //clsDiv.DivFat(sRig, x, dDdo);     /* Lettura record */
                            clsDiv.DivBol(sRig, x, dDdo);     /* Lettura record */
                        }
                        else
                            clsDiv.DivArt(sRig, x);     /* Lettura record */

                        x["div_tva"] = sTva;
                        x["div_num"] = sNum;
                        x["div_nri"] = iRig.ToString("00000");
                        iRig++;

                        j = t.Select("div_tva='" + sTva + "' AND div_num='" + sNum + "' AND div_arf='" + x["div_arf"] + "'");
                        if (j.Length == 0)
                            t.Rows.Add(x);

                        string sEan = (string)x["div_ean"];

                        if (sEan != "")
                        {
                            if (sEan.Length > 7)
                            {
                                s = sEan;
                                if (_strPar020EsclusioneEan2021davDiv == "S" && s.Length == 13 && s.Substring(7, 5) == "00000" && ((s.Substring(0, 2) == "20") || s.Substring(0, 2) == "21"))
                                    Console.WriteLine("Scarto i barcode a peso che iniziano per 20 e 21");
                                else
                                {
                                    s = sEan.Substring(0, sEan.Length - 1) + new clsCtrlCodici().FindMod10Digit(sEan.Substring(0, sEan.Length - 1));
                                    j = tDie.Select("die_arf='" + x["div_arf"] + "' AND die_ean='" + s + "'");
                                    if (j.Length == 0)
                                    {
                                        DataRow y = tDie.NewRow();
                                        y["die_num"] = sNum;
                                        y["die_for"] = strFor;
                                        y["die_arf"] = x["div_arf"];
                                        y["die_ean"] = s;
                                        tDie.Rows.Add(y);

                                        if (s != sEan)
                                            _clsFun.ErrorLog("Divulgazione DPIU", "Corretto check digit " + ((string)x["div_ean"]).Trim() + " -> " + s);
                                    }
                                }
                            }
                        }
                    }
               }

                string[] a = sNums.Split(';');

                //da controllare
                for (int i = 0; i < a.Length; i++)
                {
                    DataTable tTab = CtrlDivArt(strFil, t, strFor, a[i]);
                    if (strTip != "OFF")          //Evito controllo articoli
                    {
                        DataView v = new DataView(tTab, "tab_cod=''", "", DataViewRowState.CurrentRows);
                        if (v.Count > 0)
                        {
                            frmGesDivCtrlTab f = new frmGesDivCtrlTab();
                            f._tabTab = tTab.Copy();
                            f._tabDiv = t.Copy();
                            f._strForCod = strFor;
                            f.ShowDialog();
                            if (f._bolMdy)
                            {
                                CtrlDivArt(strFil, t, strFor, a[i]);
                            }
                        }
                    }
                }
            }
        }

        private void FillVefri(string strFil, string strFor, string strDes, string strTip, string strTfo)
        {
            DataRow x;
            string s = "";
            string sRig = "";

            using (StreamReader sr = new StreamReader(strFil, System.Text.Encoding.Default))
            {
                clsDivFietta clsDiv = new clsDivFietta();

                FileInfo fInfo = new FileInfo(strFil);

                progressBar1.Value = 0;
                progressBar1.Minimum = 0;
                progressBar1.Maximum = (int)fInfo.Length;

                DataTable tTes = (DataTable)dgv1.DataSource;

                DataTable tDie = _dasGen.Tables[TABTMPEAN];
                DataColumn[] keys = new DataColumn[2];
                keys[0] = tDie.Columns["die_arf"];
                keys[1] = tDie.Columns["die_ean"];
                tDie.PrimaryKey = keys;

                DataTable t = _dasGen.Tables[TABTMPFDR];

                string sTva = "";
                string sNdo = "";
                DateTime dDdo = DateTime.Today;
                //DateTime dDdo = _clsDef.DAYOUT;
                string sNum = "";
                string sNums = "";
                int iRig = 1;
                string sDes = "";
                DataRow[] j;

                Boolean bTes = true;

                //sNum = (tTes.Rows.Count + 1).ToString("000");
                sNum = DitNewNum();

                sTva = VARART;

                while ((sRig = sr.ReadLine()) != null)
                {
                    progressBar1.Increment(sRig.Length);
                    System.Windows.Forms.Application.DoEvents();

                    //sNum = (tTes.Rows.Count).ToString("000");
                    //sNum = DitNewNum();
                    sNum = "001";
                    if (Convert.ToDecimal(tTes.Rows.Count) > 0)
                        sNum = (string)tTes.Rows[tTes.Rows.Count - 1]["dit_num"];

                    if (strTip == "FAT" && sRig.Substring(0, 2) == "01")
                    {
                        bTes = true;
                        sTva = VARFAT;
                        sNdo = sRig.Substring(19, 6);
                        sDes = "Fattura n. " + sNdo;
                        dDdo = _clsFun.Str2Day(sRig.Substring(25, 6));
                    }

                    if (bTes)
                    {
                        //sNum = (tTes.Rows.Count + 1).ToString("000");
                        sNum = DitNewNum();

                        sNums += sNum + ";";

                        if (strTip != "FAT")
                            sDes = "Variazioni";

                        x = tTes.NewRow();
                        x["dit_num"] = sNum;
                        x["dit_for"] = strFor;
                        x["dit_des"] = sDes;
                        x["dit_fil"] = Path.GetFileName(strFil);
                        x["dit_tva"] = sTva;
                        x["dit_ndo"] = sNdo;
                        x["dit_ddo"] = dDdo;
                        x["dit_pth"] = Path.GetDirectoryName(strFil);
                        x["dit_nri"] = 1;
                        x["dit_tfo"] = strTfo;
                        x["dit_cho"] = "E";

                        tTes.Rows.Add(x);
                        iRig = 1;

                        bTes = false;
                    }
                    else if (sRig.Substring(0, 2) == "02")
                    {
                        x = t.NewRow();
                        if (strTip == "FAT")
                            clsDiv.DivFat(sRig, x, dDdo);     /* Lettura record */
                        else
                            clsDiv.DivArt(sRig, x);     /* Lettura record */

                        x["div_tva"] = sTva;
                        x["div_num"] = sNum;
                        x["div_nri"] = iRig.ToString("00000");
                        iRig++;

                        j = t.Select("div_tva='" + sTva + "' AND div_num='" + sNum + "' AND div_arf='" + x["div_arf"] + "'");
                        if (j.Length == 0)
                            t.Rows.Add(x);

                        /* Barcode */
                        //ArrayList aEan = new ArrayList();

                        //s = (string)x["div_ean"];

                        //if (s != "")
                        //{
                        //    if (strTip == "FAT")
                        //    {
                        //        //s = sRig.Substring(139);

                        //        int i = 0;
                        //        while (true)
                        //        {
                        //            if (i + 13 > s.Length)
                        //                break;

                        //            string sEan = s.Substring(i, 13);

                        //            if (sEan == new string('0', 13))
                        //                break;
                        //            else
                        //            {
                        //                sEan = Convert.ToDouble(sEan).ToString();
                        //                aEan.Add(sEan);
                        //                i += 13;
                        //            }
                        //        }

                        //    }
                        //    else
                        //        aEan.Add((string)x["div_ean"]);

                        //    //s = ((string)x["div_ean"]).Trim();

                        //    foreach (string sEan in aEan)
                        //    {

                        string sEan = (string)x["div_ean"];

                        if (sEan != "")
                        {
                            if (sEan.Length > 7)
                            {
                                s = sEan;
                                if (_strPar020EsclusioneEan2021davDiv == "S" && s.Length == 13 && s.Substring(7, 5) == "00000" && ((s.Substring(0, 2) == "20") || s.Substring(0, 2) == "21"))
                                    Console.WriteLine("Scarto i barcode a peso che iniziano per 20 e 21");
                                else
                                {
                                    s = sEan.Substring(0, sEan.Length - 1) + new clsCtrlCodici().FindMod10Digit(sEan.Substring(0, sEan.Length - 1));
                                    j = tDie.Select("die_arf='" + x["div_arf"] + "' AND die_ean='" + s + "'");
                                    if (j.Length == 0)
                                    {
                                        DataRow y = tDie.NewRow();
                                        y["die_num"] = sNum;
                                        y["die_for"] = strFor;
                                        y["die_arf"] = x["div_arf"];
                                        y["die_ean"] = s;
                                        tDie.Rows.Add(y);

                                        if (s != sEan)
                                            _clsFun.ErrorLog("Divulgazione DPIU", "Corretto check digit " + ((string)x["div_ean"]).Trim() + " -> " + s);
                                    }
                                }
                            }
                        }
                    }
                }

                string[] a = sNums.Split(';');

                //da controllare
                for (int i = 0; i < a.Length; i++)
                {
                    DataTable tTab = CtrlDivArt(strFil, t, strFor, a[i]);
                    if (strTip != "OFF")          //Evito controllo articoli
                    {
                        DataView v = new DataView(tTab, "tab_cod=''", "", DataViewRowState.CurrentRows);
                        if (v.Count > 0)
                        {
                            frmGesDivCtrlTab f = new frmGesDivCtrlTab();
                            f._tabTab = tTab.Copy();
                            f._tabDiv = t.Copy();
                            f._strForCod = strFor;
                            f.ShowDialog();
                            if (f._bolMdy)
                            {
                                CtrlDivArt(strFil, t, strFor, a[i]);
                            }
                        }
                    }
                }
            }
        }

        private void FillApOffice(string strFil, string strFor, string strDes, string strTip, string strTfo, string strForArf)
        {
            DataRow x;
            string s = "";
            string sRig = "";
            string[] a;

            using (StreamReader sr = new StreamReader(strFil, System.Text.Encoding.Default))
            {
                string sTip = strTip;

                clsDivApOffice clsDiv = new clsDivApOffice();
                clsDiv._strForArf = strForArf;

                s = "SELECT * FROM TabReparti WHERE tab_ann=0";
                DataTable tRep = _clsFun.FillTabSql("TabReparti", s, false, _strConSql);

                FileInfo fInfo = new FileInfo(strFil);

                progressBar1.Value = 0;
                progressBar1.Maximum = (int)fInfo.Length;
                progressBar1.Minimum = 0;

                DataTable tTes = (DataTable)dgv1.DataSource;

                //keys[0] = t.Columns["div_tva"];
                //keys[1] = t.Columns["div_num"];
                //keys[2] = t.Columns["div_arf"];
                //keys[3] = t.Columns["div_nri"];

                DataTable tDie = _dasGen.Tables[TABTMPEAN];
                DataColumn[] keys = new DataColumn[3];
                keys[0] = tDie.Columns["die_arf"];
                keys[1] = tDie.Columns["die_ean"];
                keys[2] = tDie.Columns["die_for"];
                tDie.PrimaryKey = keys;

                DataTable t = _dasGen.Tables[TABTMPFDR];

                string sTva = "";
                string sNdo = "";
                DateTime dDdo = _clsDef.DAYOUT;
                string sNum = "";
                string sNums = "";
                int iRig = 1;
                DataRow[] j;

                sTva = VARART;
                if (sTip == "OFF")
                    sTva = "OFF";
                //else if (sTip == "FAT")
                //    sTva = "FAT";

                string sTes = "";
                string[] aTes = s.Split('|');

                //sNum = (tTes.Rows.Count + 1).ToString("000");
                sNum = DitNewNum();

                while ((sRig = sr.ReadLine()) != null)
                {
                    progressBar1.Increment(sRig.Length);
                    System.Windows.Forms.Application.DoEvents();

                    if (sRig.Length > 10)
                    {
                        if (sRig.Substring(0, 1) == "T")
                        {
                            sTes = sRig;
                            aTes = sTes.Split('|');
                        }
                        else if (sRig.Length > 10)
                        {
                            //sNum = (tTes.Rows.Count).ToString("000");
                            //sNum = DitNewNum();
                            sNum = "001";
                            if (Convert.ToDecimal(tTes.Rows.Count) > 0)
                                sNum = (string)tTes.Rows[tTes.Rows.Count - 1]["dit_num"];

                            sNdo = sNum;
                            if (sTip == "OFF")
                            {
                                //sNdo = sRig.Substring(22, 5);

                                a = sRig.Split('|');

                                int i = Array.IndexOf(aTes, "oft_cod");
                                sNdo = a[i];
                                dDdo = DateTime.Today;

                            }
                            else if (sTip == "FAT")
                                sNdo = sRig.Substring(1, 7);

                            string sEan = "";

                            j = tTes.Select("dit_tva='" + sTva + "' AND dit_ndo='" + sNdo + "' AND dit_ddo=#" + dDdo.ToString("MM/dd/yyyy") + "#");

                            if (j.Length == 0)
                            {
                                //sNum = (tTes.Rows.Count + 1).ToString("000");
                                sNum = DitNewNum();

                                sNums += sNum + ";";

                                x = tTes.NewRow();
                                x["dit_num"] = sNum;
                                x["dit_for"] = strFor;
                                x["dit_des"] = strDes;

                                s = Path.GetFileName(strFil);
                                if (s.Length > 150)
                                    s = s.Substring(0, 150);
                                x["dit_fil"] = s;

                                x["dit_tva"] = sTva;

                                x["dit_ndo"] = sNum;    //sNdo
                                if (sTip == "OFF")
                                    x["dit_ndo"] = sNdo;
                                else if (sTip == "FAT")
                                    x["dit_ndo"] = sNdo;

                                x["dit_ddo"] = DateTime.Today;
                                if (sTva == "FAT")
                                    x["dit_ddo"] = DateTime.Today;

                                x["dit_pth"] = Path.GetDirectoryName(strFil);
                                x["dit_nri"] = 1;
                                x["dit_tfo"] = strTfo;
                                x["dit_cho"] = "E";

                                if (sTip == "OFF")
                                {
                                    a = sRig.Split('|');

                                    int i = Array.IndexOf(aTes, "oft_cod");
                                    sNdo = a[i];
                                    x["dit_ndo"] = sNdo;

                                    //s = sRig.Substring(61, 8);
                                    //s = "20" + s.Substring(6, 2) + s.Substring(3, 2) + s.Substring(0, 2);
                                    //x["dit_sii"] = _clsFun.Str2Day(s);

                                    //s = sRig.Substring(6, 8);
                                    //s = "20" + s.Substring(6, 2) + s.Substring(3, 2) + s.Substring(0, 2);
                                    //x["dit_sif"] = _clsFun.Str2Day(s);

                                    //s = sRig.Substring(6, 8);
                                    //s = "20" + s.Substring(6, 2) + s.Substring(3, 2) + s.Substring(0, 2);

                                    i = Array.IndexOf(aTes, "oft_dti");
                                    s = a[i];
                                    x["dit_soi"] = _clsFun.Str2Day(s);

                                    i = Array.IndexOf(aTes, "oft_dtf");
                                    s = a[i];
                                    x["dit_sof"] = _clsFun.Str2Day(s);

                                    i = Array.IndexOf(aTes, "oft_des");
                                    s = a[i];
                                    x["dit_des"] = s.Trim();

                                }

                                tTes.Rows.Add(x);
                                iRig = 1;
                                dDdo = (DateTime)x["dit_ddo"];
                            }
                            else
                            {
                                sNum = (string)j[0]["dit_num"];
                            }

                            s = sRig.Substring(2, 7);
                            x = t.NewRow();
                            j = t.Select("div_tva='" + sTva + "' AND div_num='" + sNum + "' AND div_arf='" + s + "'");
                            if (j.Length > 0)
                                x = j[0];
                            //x = t.NewRow();

                            if (sTip == "VAR")
                                sEan = clsDiv.DivVar(sTes, sRig, x);     /* Lettura record */
                            else if (sTip == "OFF")
                                sEan = clsDiv.DivVarOff(sTes, sRig, x);     /* Lettura record */
                            //else if (sTip == "VAR")
                            //    clsDiv.DivArt(sRig, x);     /* Lettura record */
                            //else if (sTip == "OFF")
                            //    clsDiv.DivOff(sRig, x);     /* Lettura record */
                            //else if (sTip == "FAT")
                            //    clsDiv.DivFat(sRig, x);     /* Lettura record */
                            x["div_tva"] = sTva;
                            x["div_num"] = sNum;
                            x["div_nri"] = iRig.ToString("00000");
                            iRig++;

                            if (s == "0155310")
                                Console.WriteLine("aaaaa");

                            j = t.Select("div_tva='" + sTva + "' AND div_num='" + sNum + "' AND div_arf='" + (string)x["div_arf"] + "'");
                            if (j.Length == 0)
                            {
                                j = tRep.Select("tab_cod='" + x["div_rep"] + "'");
                                if (j.Length > 0)
                                    t.Rows.Add(x);
                            }
                            /* Barcode */
                            //s = ((string)x["ean_ean"]).Trim();
                            if (sEan != "")
                            {
                                string[] aa = sEan.Split('?');

                                foreach (string sss in aa)
                                {
                                    string[] ss = sss.Split(';');

                                    if (ss.Length > 5)
                                    {
                                        string sEEan = ss[0];
                                        decimal dEQta = Convert.ToDecimal(ss[1]);
                                        decimal dEPrv = Convert.ToDecimal(ss[2]);
                                        Boolean bEBil = false;
                                        if (ss[3] == "1")
                                            bEBil = true;
                                        Boolean bEEcp = false;
                                        if (ss[4] == "1")
                                            bEEcp = true;
                                        Boolean bEAnn = false;
                                        if (ss[5] == "1")
                                            bEAnn = true;

                                        if (sEEan.Length > 7)
                                        {
                                            s = sEEan;
                                            if (_strPar020EsclusioneEan2021davDiv == "S" && s.Length == 13 && s.Substring(7, 5) == "00000" && ((s.Substring(0, 2) == "20") || s.Substring(0, 2) == "21") && !bEBil)
                                                Console.WriteLine("Scarto i barcode a peso che iniziano per 20 e 21 a meno che non sia attivo in bilancia");
                                            else
                                            {
                                                if (sEEan.Length == 13 && sEEan.Substring(0, 1) == "2")
                                                    s = sEEan;
                                                else
                                                    s = sEEan.Substring(0, sEEan.Length - 1) + new clsCtrlCodici().FindMod10Digit(sEEan.Substring(0, sEEan.Length - 1));

                                                DataRow y = tDie.NewRow();
                                                j = tDie.Select("die_arf='" + x["div_arf"] + "' AND die_ean='" + s + "'");
                                                if (j.Length > 0)
                                                    y = j[0];

                                                y["die_num"] = sNum;
                                                y["die_for"] = strFor;
                                                y["die_arf"] = x["div_arf"];
                                                y["die_ean"] = s;

                                                y["die_qta"] = dEQta;
                                                y["die_prv"] = dEPrv;
                                                y["die_ecp"] = bEEcp;
                                                y["die_bil"] = bEBil;
                                                y["die_ann"] = bEAnn;

                                                j = tDie.Select("die_arf='" + x["div_arf"] + "' AND die_ean='" + s + "'");
                                                if (j.Length == 0)
                                                    tDie.Rows.Add(y);

                                                //if (s != ((string)x["div_ean"]).Trim())
                                                if (s != sEEan)
                                                    _clsFun.ErrorLog("Divulgazione DPIU", "Corretto check digit " + sEan + " -> " + s);

                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                }

                a = sNums.Split(';');
                for (int i = 0; i < a.Length; i++)
                {
                    //DataTable tTab = CtrlDivArt(strFil, t, strFor, a[i]);
                    DataTable tTab = CtrlDivArt("DIVAPOFFICE", t, strFor, a[i]);
                    //if (sTip != "OFF")          //Evito controllo articoli
                    //{
                    //    DataView v = new DataView(tTab, "tab_cod=''", "", DataViewRowState.CurrentRows);
                    //    if (v.Count > 0)
                    //    {
                    //        frmGesDivCtrlTab f = new frmGesDivCtrlTab();
                    //        f._tabTab = tTab.Copy();
                    //        f._tabDiv = t.Copy();
                    //        f._strForCod = strFor;
                    //        f.ShowDialog();
                    //        if (f._bolMdy)
                    //        {
                    //            CtrlDivArt(strFil, t, strFor, a[i]);
                    //        }
                    //    }
                    //}
                }
            }
        }

        /*** APOFFICE fine ***/

        private void FillDadoArt(string strFil, string strFor, string strDes)
        {
            DataRow x;
            string s = "";
            string sRig = "";

            using (StreamReader sr = new StreamReader(strFil, System.Text.Encoding.Default))
            {
                clsDivDado clsDiv = new clsDivDado();

                FileInfo fInfo = new FileInfo(strFil);

                progressBar1.Value = 0;
                progressBar1.Maximum = (int)fInfo.Length;
                progressBar1.Minimum = 0;

                DataTable tTes = (DataTable)dgv1.DataSource;

                DataTable tDie = _dasGen.Tables[TABTMPEAN];
                DataColumn[] keys = new DataColumn[2];
                keys[0] = tDie.Columns["die_arf"];
                keys[1] = tDie.Columns["die_ean"];
                tDie.PrimaryKey = keys;

                DataTable t = _dasGen.Tables[TABTMPFDR];

                string sTva = "ART";
                string sNdo = "";
                DateTime dDdo = _clsDef.DAYOUT;
                string sNum = "";
                string sNums = "";
                int iRig = 1;
                DataRow[] j;
                Boolean b = true;

                while ((sRig = sr.ReadLine()) != null)
                {
                    progressBar1.Increment(sRig.Length);
                    System.Windows.Forms.Application.DoEvents();

                    if (sRig.Length > 10 && iRig > 1)
                    {
                        if (b)
                        {
                            b = false;
                            //s = sRig.Substring(184, 8);
                            //dDdo = new DateTime(Convert.ToInt32(s.Substring(0, 4)), Convert.ToInt32(s.Substring(4, 2)), Convert.ToInt32(s.Substring(6, 2)));
                            dDdo = DateTime.Today;
                            j = tTes.Select("dit_tva='" + sTva + "' AND dit_ndo='" + sNdo + "' AND dit_ddo=#" + dDdo.ToString("MM/dd/yyyy") + "#");
                            if (j.Length == 0)
                            {
                                //sNum = (tTes.Rows.Count + 1).ToString("000");
                                sNum = DitNewNum();

                                sNums += sNum + ";";
                                x = tTes.NewRow();
                                x["dit_num"] = sNum;
                                x["dit_for"] = strFor;
                                x["dit_des"] = strDes;
                                x["dit_fil"] = Path.GetFileName(strFil);
                                x["dit_tva"] = sTva;
                                x["dit_ndo"] = sNdo;
                                x["dit_ddo"] = dDdo;
                                x["dit_pth"] = Path.GetDirectoryName(strFil);
                                x["dit_nri"] = 1; 
                                x["dit_cho"] = "E";

                                tTes.Rows.Add(x);
                                iRig = 1;
                                dDdo = (DateTime)x["dit_ddo"];
                            }
                        }

                        x = t.NewRow();
                        clsDiv.DivArt(sRig, x);     /***/
                        x["div_tva"] = sTva;
                        x["div_num"] = sNum;
                        x["div_nri"] = iRig.ToString("00000");

                        j = t.Select("div_tva='" + sTva + "' AND div_num='" + sNum + "' AND div_arf='" + x["div_arf"] + "'");
                        if (j.Length == 0)
                            t.Rows.Add(x);

                        if ((string)x["div_arf"] == "2534801")
                            Console.WriteLine("aaaaaaaaa");

                        s = ((string)x["div_ean"]).Trim();
                        if (s != "" && s.Length > 7 && _clsFun.Numerico(s))
                        {
                            s = s.Substring(0, s.Length - 1) + new clsCtrlCodici().FindMod10Digit(s.Substring(0, s.Length - 1));
                            j = tDie.Select("die_arf='" + x["div_arf"] + "' AND die_ean='" + s + "'");
                            if (j.Length == 0)
                            {
                                if (s == "0178")
                                    Console.WriteLine("aaaaaaaaaa");

                                DataRow y = tDie.NewRow();
                                y["die_num"] = sNum;
                                y["die_for"] = strFor;
                                y["die_arf"] = x["div_arf"];
                                y["die_ean"] = s;
                                tDie.Rows.Add(y);

                                if (s != ((string)x["div_ean"]).Trim())
                                    _clsFun.ErrorLog("Divulgazione DADO", "Corretto check digit " + ((string)x["div_ean"]).Trim() + " -> " + s);
                            }
                        }
                    }

                    iRig++;
                    //}
                    //break;
                }

                string[] a = sNums.Split(';');
                for (int i = 0; i < a.Length; i++)
                {
                    DataTable tTab = CtrlDivArt(strFil, t, strFor, a[i]);
                    DataView v = new DataView(tTab, "tab_cod=''", "", DataViewRowState.CurrentRows);
                    if (v.Count > 0)
                    {
                        frmGesDivCtrlTab f = new frmGesDivCtrlTab();
                        f._tabTab = tTab.Copy();
                        f._strForCod = strFor;
                        f.ShowDialog();
                        if (f._bolMdy)
                        {
                            CtrlDivArt(strFil, t, strFor, a[i]);
                        }
                    }
                }
            }
        }

        private void FillSMA(string strFil, string strFor, string strDes, string strNeg)
        {
            clsDivSMA clsDiv = new clsDivSMA();

            DataRow x;
            string s = "";
            string sRig = "";

            s = Path.GetDirectoryName(strFil) + "\\Listino.xls" ;
            lblDiv.Text = "Caricamento listino SMA";
            DataTable tSma = clsDiv.Xls2Tab("LIS", s, progressBar1);

            using (StreamReader sr = new StreamReader(strFil, System.Text.Encoding.Default))
            {
                FileInfo fInfo = new FileInfo(strFil);

                progressBar1.Value = 0;
                progressBar1.Maximum = (int)fInfo.Length;
                progressBar1.Minimum = 0;

                DataTable tTes = (DataTable)dgv1.DataSource;

                DataTable tDie = _dasGen.Tables[TABTMPEAN];
                DataColumn[] keys = new DataColumn[2];
                keys[0] = tDie.Columns["die_arf"];
                keys[1] = tDie.Columns["die_ean"];
                tDie.PrimaryKey = keys;

                DataTable t = _dasGen.Tables[TABTMPFDR];

                string sTva = "ART";
                string sNdo = "";
                DateTime dDdo = _clsDef.DAYOUT;
                string sNum = "";
                string sNums = "";
                int iRig = 1;
                DataRow[] j;
                Boolean b = true;
                string sMsg = "";
                string sNeg = "";

                string[] aNeg = strNeg.Split(',');
                if (aNeg.Length > 1)
                {
                    s = Path.GetFileName(strFil);

                    if (s.Length > 8)
                    {

                        s = s.Substring(4, 4);
                        for (int ii = 0; ii < aNeg.Length; ii++)
                        {
                            if (aNeg[ii] == s)
                            {
                                sNeg = (ii + 1).ToString("000");
                                break;
                            }
                        }
                    }
                }
                if(sNeg == "")
                    sNeg = "001";

                while ((sRig = sr.ReadLine()) != null)
                {
                    progressBar1.Increment(sRig.Length);
                    System.Windows.Forms.Application.DoEvents();

                    if (sRig.Length > 10 && sRig.Substring(0, 2) == "04")
                        Console.WriteLine("XXXXX");

                    if (sRig.Length > 10 && (sRig.Substring(0, 2) == "01" || sRig.Substring(0, 2) == "04" || sRig.Substring(0, 2) == "05"))
                    {
                        clsDiv.DivSmaAna(sRig, tSma);     /***/
                    }
                    else if (sRig.Length > 160 && sRig.Substring(0,2) == "21")
                    {
                        sTva = VARFAT;
                        dDdo = _clsFun.Str2Day(sRig.Substring(105,10).Replace("-",""));
                        //sNdo = sRig.Substring(115, 6);
                        sNdo = sRig.Substring(99, 6);

                        //if (b)
                        //{
                        //    b = false;
                        //s = sRig.Substring(184, 8);
                        //dDdo = new DateTime(Convert.ToInt32(s.Substring(0, 4)), Convert.ToInt32(s.Substring(4, 2)), Convert.ToInt32(s.Substring(6, 2)));
                        //dDdo = DateTime.Today;
                        j = tTes.Select("dit_tva='" + sTva + "' AND dit_ndo='" + sNdo + "' AND dit_ddo=#" + dDdo.ToString("MM/dd/yyyy") + "#");
                        if (j.Length == 0)
                        {
                            //sNum = (tTes.Rows.Count + 1).ToString("000");
                            sNum = DitNewNum();

                            sNums += sNum + ";";
                            x = tTes.NewRow();
                            x["dit_num"] = sNum;
                            x["dit_for"] = strFor;
                            x["dit_des"] = strDes;
                            x["dit_fil"] = Path.GetFileName(strFil);
                            x["dit_tva"] = sTva;
                            x["dit_ndo"] = sNdo;
                            x["dit_ddo"] = dDdo;
                            x["dit_pth"] = Path.GetDirectoryName(strFil);
                            x["dit_nri"] = 1;
                            x["dit_cho"] = "E";
                            x["dit_neg"] = sNeg;

                            tTes.Rows.Add(x);
                            iRig = 1;
                            dDdo = (DateTime)x["dit_ddo"];
                        }
                        //}

                        x = t.NewRow();
                        clsDiv.DivFat(sRig, x, tSma, ref sMsg);     /***/
                        x["div_tva"] = sTva;
                        x["div_num"] = sNum;
                        x["div_nri"] = iRig.ToString("00000");

                        if (!DBNull.Value.Equals(x["div_ean"]) && ((string)x["div_ean"]).Trim() != "")
                        {
                            j = t.Select("div_tva='" + sTva + "' AND div_num='" + sNum + "' AND div_arf='" + x["div_arf"] + "'");
                            if (j.Length == 0)
                                t.Rows.Add(x);

                            //if ((string)x["div_arf"] == "2534801")
                            //    Console.WriteLine("aaaaaaaaa");

                            s = ((string)x["div_ean"]).Trim();
                            if (s != "" && s.Length > 7 && _clsFun.Numerico(s))
                            {
                                s = s.Substring(0, s.Length - 1) + new clsCtrlCodici().FindMod10Digit(s.Substring(0, s.Length - 1));
                                j = tDie.Select("die_arf='" + x["div_arf"] + "' AND die_ean='" + s + "'");
                                if (j.Length == 0)
                                {
                                    if (s == "0178")
                                        Console.WriteLine("aaaaaaaaaa");

                                    DataRow y = tDie.NewRow();
                                    y["die_num"] = sNum;
                                    y["die_for"] = strFor;
                                    y["die_arf"] = x["div_arf"];
                                    y["die_ean"] = s;
                                    tDie.Rows.Add(y);

                                    if (s != ((string)x["div_ean"]).Trim())
                                        _clsFun.ErrorLog("Divulgazione DADO", "Corretto check digit " + ((string)x["div_ean"]).Trim() + " -> " + s);
                                }
                            }
                        }
                    }

                    iRig++;
                    //}
                    //break;
                }

                string[] a = sNums.Split(';');
                for (int i = 0; i < a.Length; i++)
                {
                    DataTable tTab = CtrlDivArt(strFil, t, strFor, a[i]);
                    DataView v = new DataView(tTab, "tab_cod=''", "", DataViewRowState.CurrentRows);
                    if (v.Count > 0)
                    {
                        frmGesDivCtrlTab f = new frmGesDivCtrlTab();
                        f._tabTab = tTab.Copy();
                        f._strForCod = strFor;
                        f.ShowDialog();
                        if (f._bolMdy)
                        {
                            CtrlDivArt(strFil, t, strFor, a[i]);
                        }
                    }
                }

                if (sMsg != "")
                    MessageBox.Show(sMsg, "ARTICOLI NON TROVATI SUL LISTINO", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
        }

        private void FillApDivVar(string strFil, string strFor, string strDes)
        {
            DataRow x;
            string s = "";
            string sRig = "";

            using (StreamReader sr = new StreamReader(strFil, System.Text.Encoding.Default))
            {
                clsDivApDiv clsDiv = new clsDivApDiv();

                FileInfo fInfo = new FileInfo(strFil);

                progressBar1.Value = 0;
                progressBar1.Maximum = (int)fInfo.Length;
                progressBar1.Minimum = 0;

                DataTable tTes = (DataTable)dgv1.DataSource;

                DataTable tDie = _dasGen.Tables[TABTMPEAN];
                DataColumn[] keys = new DataColumn[2];
                keys[0] = tDie.Columns["die_arf"];
                keys[1] = tDie.Columns["die_ean"];
                tDie.PrimaryKey = keys;

                DataTable t = _dasGen.Tables[TABTMPFDR];

                string sTva = "";
                string sNdo = "";
                DateTime dDdo = _clsDef.DAYOUT;
                string sNum = "";
                string sNums = "";
                int iRig = 1;
                DataRow[] j;

                sTva = VARART;
                //if (Path.GetFileName(strFil).Substring(0, 7) == "KONZMOV")
                //{
                //    if (strFil.ToLower().IndexOf("impianto") == 0)
                //        sTva = VARFAT;
                //}

                while ((sRig = sr.ReadLine()) != null)
                {
                    progressBar1.Increment(sRig.Length);
                    System.Windows.Forms.Application.DoEvents();

                    sNdo = ""; // sRig.Substring(6, 6);

                    j = tTes.Select("dit_tva='" + sTva + "' AND dit_ndo='" + sNdo + "' AND dit_ddo=#" + dDdo.ToString("MM/dd/yyyy") + "#");
                    if (j.Length == 0)
                    {
                        //sNum = (tTes.Rows.Count + 1).ToString("000");
                        sNum = DitNewNum();

                        sNums += sNum + ";";
                        x = tTes.NewRow();
                        x["dit_num"] = sNum;
                        x["dit_for"] = strFor;
                        x["dit_des"] = strDes;
                        x["dit_fil"] = Path.GetFileName(strFil);
                        x["dit_tva"] = sTva;
                        x["dit_ndo"] = sNdo;
                        x["dit_ddo"] = DateTime.Today; // new DateTime(Convert.ToInt32(sRig.Substring(16, 4)), Convert.ToInt32(sRig.Substring(14, 2)), Convert.ToInt32(sRig.Substring(12, 2)));
                        x["dit_pth"] = Path.GetDirectoryName(strFil);
                        x["dit_nri"] = 1;
                        x["dit_cho"] = "E";

                        tTes.Rows.Add(x);
                        iRig = 1;
                        dDdo = (DateTime)x["dit_ddo"];
                    }
                    else
                    {
                        sNum = (string)j[0]["dit_num"];
                        j[0]["dit_nri"] = (decimal)j[0]["dit_nri"] + 1;
                        iRig = Convert.ToInt32(j[0]["dit_nri"]);
                    }

                    string[] aRig = sRig.Split(';');

                    Boolean b = true;

                    s = aRig[0];

                    if (s.Length > 2 && s.Substring(0, 2).ToLower() == "sp")
                        s = s.Substring(2);

                    aRig[0] = s;

                    if (s.Length == 0 || !_clsFun.Numerico(s))
                        b = false;

                    if (b)
                    {
                        x = t.NewRow();
                        clsDiv.DivArt(aRig, x);     /***/
                        x["div_tva"] = sTva;
                        x["div_num"] = sNum;
                        x["div_nri"] = iRig.ToString("00000");

                        j = t.Select("div_tva='" + sTva + "' AND div_num='" + sNum + "' AND div_arf='" + x["div_arf"] + "'");
                        if (j.Length == 0)
                            t.Rows.Add(x);

                        //if ((string)x["div_arf"] == "026200")
                        //    Console.WriteLine("aaaaaaaaa");

                        s = ((string)x["div_ean"]).Trim();
                        if (s != "")
                        {
                            if(s.Length == 13 && s.Substring(0,1) != "2")
                                s = s.Substring(0, s.Length - 1) + new clsCtrlCodici().FindMod10Digit(s.Substring(0, s.Length - 1));
                            j = tDie.Select("die_arf='" + x["div_arf"] + "' AND die_ean='" + s + "'");
                            if (j.Length == 0)
                            {
                                DataRow y = tDie.NewRow();
                                y["die_num"] = sNum;
                                y["die_for"] = strFor;
                                y["die_arf"] = x["div_arf"];
                                y["die_ean"] = s;
                                tDie.Rows.Add(y);

                                if (s != ((string)x["div_ean"]).Trim())
                                    _clsFun.ErrorLog("Divulgazione TERRON", "Corretto check digit " + ((string)x["div_ean"]).Trim() + " -> " + s);
                            }
                        }
                    }
                    //break;
                }

                string[] a = sNums.Split(';');
                for (int i = 0; i < a.Length; i++)
                {

                    DataTable tTab = CtrlDivArt(strFil, t, strFor, a[i]);
                    DataView v = new DataView(tTab, "tab_cod=''", "", DataViewRowState.CurrentRows);
                    if (v.Count > 0)
                    {
                        frmGesDivCtrlTab f = new frmGesDivCtrlTab();
                        f._tabTab = tTab.Copy();
                        f._strForCod = strFor;
                        f.ShowDialog();
                        if (f._bolMdy)
                        {
                            CtrlDivArt(strFil, t, strFor, a[i]);
                        }
                    }
                }
            }
        }

        private void FillTmpDiv()
        {
            string sTab = TABTMPDIV.ToUpper();
            string sTva = VARART;

            string s = "SELECT * FROM TmpDiv WHERE tmp_tip='ADV' AND tmp_sta='S'";
            DataTable tDiv = _clsFun.FillTabSql("TmpDiv", s, false, _strConSql);
            if (tDiv.Rows.Count > 0)
            {
                DataRow[] j;
                DataRow x;
                DataTable t = _dasGen.Tables[TABTMPFDR];

                progressBar1.Value = 0;
                progressBar1.Maximum = t.Rows.Count;
                progressBar1.Minimum = 0;

                DataTable tTes = (DataTable)dgv1.DataSource;

                //string sNum = (tTes.Rows.Count + 1).ToString("000");
                string sNum = DitNewNum();

                x = tTes.NewRow();
                x["dit_num"] = sNum;
                x["dit_for"] = "";
                x["dit_des"] = "Divulgazione interna";
                x["dit_fil"] = sTab;
                x["dit_tva"] = sTva;
                x["dit_ndo"] = "";
                x["dit_ddo"] = DateTime.Today.ToShortDateString(); // new DateTime(Convert.ToInt32(sRig.Substring(16, 4)), Convert.ToInt32(sRig.Substring(14, 2)), Convert.ToInt32(sRig.Substring(12, 2)));
                x["dit_pth"] = "";
                x["dit_nri"] = 1;
                x["dit_cho"] = "E";

                tTes.Rows.Add(x);

                int iRig = 1;
                DateTime dDdo = (DateTime)x["dit_ddo"];

                string sTime = "";

                foreach(DataRow y in tDiv.Rows)
                {
                    string sRig = (string)y["tmp_tmp"];

                    string[] aRig = sRig.Split(';');

                    string sArt = aRig[2].Trim();
                    string sArd = aRig[3].Trim();
                    s = aRig[7].Replace(".",",");

                    decimal dPrv = 0;
                    if (s != "" && _clsFun.Numerico(s))
                    {
                        try
                        {
                            dPrv = Convert.ToDecimal(s);
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }

                    }
                    j = t.Select("div_tva='" + sTva + "' AND div_num='" + sNum + "' AND div_art='" + sArt + "'");
                    if (j.Length == 0)
                    {
                        x = t.NewRow();
                        x["div_tva"] = sTva;
                        x["div_num"] = sNum;
                        x["div_dva"] = DateTime.Today.ToShortDateString();
                        x["div_nri"] = iRig.ToString("00000");
                        x["div_art"] = sArt;
                        x["div_ard"] = sArd;
                        x["div_arf"] = "";
                        x["var_prp"] = dPrv;
                        x["div_eti"] = "S";
                        t.Rows.Add(x);
                    }
                    else
                        j[0]["var_prp"] = dPrv;

                    iRig++;

                    if (sTime == "" || Convert.ToDouble(sTime) < Convert.ToDouble((string)y["tmp_day"]))
                        sTime = (string)y["tmp_day"];
                }

                j = tTes.Select("dit_num='" + sNum + "'");
                if (j.Length > 0)
                    j[0]["dit_pth"] = sTime;

                CtrlDivArt(sTab, t, "", sNum);
            }
        }

        private void elaboraToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Color bc = lblDiv.BackColor;
            lblDiv.BackColor = Color.LightCoral;
            lblDiv.Text = "Aggiornamento in corso ...";
            System.Windows.Forms.Application.DoEvents();

            //elaboraToolStripMenuItem.Enabled = false;
            //elaboraToolStripMenuItem.Enabled = true;
            DivAggiorna();
            DivFatture();
            Pulizia();

            lblDiv.BackColor = bc;
            lblDiv.Text = "";

            MessageBox.Show("FINE ELABORAZIONE");
            this.Close();
            this.Dispose();
        }

        private void dgv1_CurrentCellChanged(object sender, EventArgs e)
        {
            FillRighe();
        }

        private void FillRighe()
        {
            try
            {
                string sTva = "";
                string sNum = "";
                CurrencyManager cm = dgv1.BindingContext[dgv1.DataSource, dgv1.DataMember] as CurrencyManager;
                if (cm.Position >= 0)
                {
                    DataRowView r = cm.List[dgv1.CurrentRow.Index] as DataRowView;
                    DataRow x = r.Row;
                    sTva = (string)x["dit_tva"];
                    sNum = (string)x["dit_num"];
                    dgv2.DataSource = new DataView(_dasGen.Tables[TABTMPFDR], "div_tva='" + sTva + "' AND div_num='" + sNum + "'", "div_nri", DataViewRowState.CurrentRows);
                    lblCnt.Text = dgv2.RowCount.ToString();
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }

        private DataTable CtrlDivArt(string strFil, DataTable tabDiv, string strFor, string strNum)
        {
            string sForRif = _clsFun.ParGet(clsDefine.enuParametri.ParForDefault, _strConSql);
            if (sForRif.Length > 5)
                sForRif = sForRif.Substring(0, 5);
            
            DataRow[] j;
            string sMsg = "";
            decimal d = 0;
            Boolean b = true;

            string sRep = "";

            string p = "AnaConfigurazione";
            string s = "SELECT * FROM " + p;
            DataTable tCnf = _clsFun.FillTabSql(p, s, false, _strConSql);
            if (tCnf.Rows.Count > 0)
                sRep = (string)tCnf.Rows[0]["cnf_rfo"];
            if (sRep.Trim() == "")
                sRep = "099";

            s = "SELECT ";
            s += "GesOffArticoli.ofa_yea, ";
            s += "GesOffArticoli.ofa_cod, ";
            s += "GesOffArticoli.ofa_art, ";
            s += "GesOffArticoli.ofa_tip, ";
            s += "GesOffArticoli.ofa_cos, ";
            s += "GesOffArticoli.ofa_prv, ";
            s += "GesOffArticoli.ofa_val, ";
            s += "GesOffArticoli.ofa_xem, ";
            s += "GesOffArticoli.ofa_xen, ";
            s += "GesOffArticoli.ofa_ann, ";
            s += "GesOffArticoli.ofa_mix, ";
            s += "GesOffTestate.oft_dti, ";
            s += "GesOffTestate.oft_dtf, ";
            s += "TabOfferte.tab_des ";
            s += "FROM GesOffArticoli ";
            s += "INNER JOIN GesOffTestate ON GesOffArticoli.ofa_yea = GesOffTestate.oft_yea AND GesOffArticoli.ofa_cod = GesOffTestate.oft_cod ";
            s += "INNER JOIN TabOfferte ON GesOffArticoli.ofa_tip = TabOfferte.tab_cod ";
            s += "WHERE (GesOffTestate.oft_dti > " + _clsFun.DaySql(DateTime.Now.AddDays(-45)) + ") AND GesOffArticoli.ofa_ann=0 ";
            s += "ORDER BY GesOffTestate.oft_dtf DESC";
            DataTable tOff = _clsFun.FillTabSql(TABANAOFF, s, false, _strConSql);
            DataColumn[] keys = new DataColumn[3];
            keys[0] = tOff.Columns["ofa_art"];
            keys[1] = tOff.Columns["oft_dti"];
            keys[2] = tOff.Columns["ofa_cod"];
            tOff.PrimaryKey = keys;

            s = "SELECT ean_ean, ean_art FROM AnaBarcode WHERE ean_ann=0";
            DataTable tEan = _clsFun.FillTabSql(TABANAEAN, s, false, _strConSql);
            keys = new DataColumn[2];
            keys[0] = tEan.Columns["ean_ean"];
            keys[1] = tEan.Columns["ean_art"];
            tEan.PrimaryKey = keys;

            DataTable tAcqFop = new DataTable();

            if(strFor != sForRif)
            {
                s = "SELECT lia_art FROM GesLisAcquisto ";
                s += "WHERE lia_for='" + sForRif + "' AND lia_ann=0 ";
                s += "GROUP BY lia_art ";
                tAcqFop = _clsFun.FillTabSql(TABLISACQ, s, false, _strConSql);
                keys = new DataColumn[5];
                keys[0] = tAcqFop.Columns["lia_art"];
                tAcqFop.PrimaryKey = keys;

                //tAcqFop = _clsFun.FillTabSql(TABLISACQ, s, false, _strConSql);
            }

            DataTable tAcq = new DataTable();

            if (strFil == "TMPDIV")
            {
                s = "SELECT * FROM AnaArticoli";
                tAcq = _clsFun.FillTabSql(TABANAART, s, false, _strConSql);
                keys = new DataColumn[1];
                keys[0] = tAcq.Columns["art_cod"];
                tAcq.PrimaryKey = keys;
            }
            else if (strFil == "DIVAPOFFICE")
            {
                s = "SELECT ";
                s += "GesLisAcquisto_1.lia_tip, ";
                s += "GesLisAcquisto_1.lia_art, ";
                s += "GesLisAcquisto_1.lia_cos, ";
                s += "GesLisAcquisto_1.lia_dti, ";
                s += "GesLisAcquisto_1.lia_dtf, ";
                s += "GesLisAcquisto_1.lia_arf, ";
                s += "GesLisAcquisto_1.lia_ann, ";
                s += "GesLisAcquisto_1.lia_pxc, ";
                s += "AnaArticoli.art_des, ";
                s += "AnaArticoli.art_iva, ";
                s += "AnaArticoli.art_sfr, ";
                s += "AnaArticoli.art_tar, ";
                s += "AnaArticoli.art_umi, ";
                s += "AnaArticoli.art_tgr, ";
                s += "AnaArticoli.art_rep, ";
                s += "AnaArticoli.art_ec1, ";
                s += "AnaArticoli.art_ec2, ";
                s += "AnaArticoli.art_ec3, ";
                s += "AnaArticoli.art_sta ";
                s += "FROM ";
                s += "(SELECT lia_tip, lia_for, MAX(lia_dti) AS lia_dti, lia_art ";
                s += " FROM GesLisAcquisto ";
                s += " GROUP BY lia_tip, lia_for, lia_art) AS A ";
                s += "LEFT OUTER JOIN ";
                s += "GesLisAcquisto AS GesLisAcquisto_1 ON A.lia_tip = GesLisAcquisto_1.lia_tip AND A.lia_for = GesLisAcquisto_1.lia_for AND A.lia_art = GesLisAcquisto_1.lia_art LEFT OUTER JOIN ";
                s += "AnaArticoli ON AnaArticoli.art_cod = A.lia_art ";
                tAcq = _clsFun.FillTabSql(TABANAART, s, false, _strConSql);
                keys = new DataColumn[1];
                keys[0] = tAcq.Columns["art_cod"];
                tAcq.PrimaryKey = keys;
            }
            else
            {
                s = "SELECT ";
                s += "lia_tip, ";
                s += "lia_for, ";
                s += "lia_arf, ";
                s += "lia_art, ";
                s += "lia_pxc, ";
                s += "lia_cos, ";
                s += "lia_dti, ";
                s += "lia_dtf, ";
                s += "lia_ann, ";
                s += "art_des, ";
                s += "art_sta, ";
                s += "art_iva, ";
                s += "art_sfr, ";
                s += "art_tar, ";
                s += "art_umi, ";
                s += "art_tgr, ";
                s += "art_rep, ";
                s += "art_ec1, ";
                s += "art_ec2, ";
                s += "art_ec3 ";
                s += "FROM GesLisAcquisto ";
                s += "LEFT OUTER JOIN AnaArticoli ON GesLisAcquisto.lia_art = AnaArticoli.art_cod ";
                s += "WHERE lia_ann = 0 AND lia_for='" + strFor + "' ORDER BY lia_dti DESC";
                tAcq = _clsFun.FillTabSql(TABLISACQ, s, false, _strConSql);
                keys = new DataColumn[5];
                keys[0] = tAcq.Columns["lia_tip"];
                keys[1] = tAcq.Columns["lia_for"];
                keys[2] = tAcq.Columns["lia_arf"];
                keys[3] = tAcq.Columns["lia_dti"];
                keys[4] = tAcq.Columns["lia_dtf"];
                tAcq.PrimaryKey = keys;
            }
            //s = "SELECT liv_art, liv_lis, liv_prv, liv_dti, liv_dtf, liv_ann FROM GesLisVendita WHERE liv_ann=0 ORDER BY liv_art, liv_dti DESC";
            s = "SELECT liv_art, liv_lis, liv_prv, liv_dti, liv_dtf, liv_ann FROM GesLisVendita ORDER BY liv_art, liv_dti DESC";
            DataTable tLiv = _clsFun.FillTabSql(TABLISVEN, s, false, _strConSql);
            keys = new DataColumn[4];
            keys[0] = tLiv.Columns["liv_lis"];
            keys[1] = tLiv.Columns["liv_art"];
            keys[2] = tLiv.Columns["liv_dti"];
            keys[3] = tLiv.Columns["liv_dtf"];
            tLiv.PrimaryKey = keys;

            s = "SELECT * FROM TabTabFornitori WHERE tab_for='" + strFor + "'";
            DataTable tTab = _clsFun.FillTabSql(TABTABFOR, s, false, _strConSql);
            tTab.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "TabArt",
                Caption = "Art",
                MaxLength = 70,
                ReadOnly = false,
                DefaultValue = (String)""
            });
            tTab.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "TabMsg",
                Caption = "Art",
                MaxLength = 70,
                ReadOnly = false,
                DefaultValue = (String)""
            });
            keys = new DataColumn[3];
            keys[0] = tTab.Columns["tab_for"];
            keys[1] = tTab.Columns["tab_tip"];
            keys[2] = tTab.Columns["tab_cof"];
            tTab.PrimaryKey = keys;

            DataTable tTmp = tabDiv.Copy();
            keys = new DataColumn[4];
            keys[0] = tTmp.Columns["div_tva"];
            keys[1] = tTmp.Columns["div_num"];
            keys[2] = tTmp.Columns["div_arf"];
            keys[3] = tTmp.Columns["div_nri"];
            tTmp.PrimaryKey = keys;

            p = "TabIva";
            s = "SELECT * FROM " + p;
            DataTable tIva = _clsFun.FillTabSql(p, s, false, _strConSql);

            p = "TabUmi";
            s = "SELECT * FROM " + p;
            DataTable tUmi = _clsFun.FillTabSql(p, s, false, _strConSql);

            p = "TabTipoGrammatura";
            s = "SELECT * FROM " + p;
            DataTable tTgr = _clsFun.FillTabSql(p, s, false, _strConSql);

            p = "TabReparti";
            s = "SELECT * FROM " + p;
            DataTable tRep = _clsFun.FillTabSql(p, s, false, _strConSql);

            p = "TabEcrLv1";
            s = "SELECT * FROM " + p;
            DataTable tEc1 = _clsFun.FillTabSql(p, s, false, _strConSql);

            p = "TabEcrLv2";
            s = "SELECT * FROM " + p;
            DataTable tEc2 = _clsFun.FillTabSql(p, s, false, _strConSql);

            p = "TabEcrLv3";
            s = "SELECT * FROM " + p;
            DataTable tEc3 = _clsFun.FillTabSql(p, s, false, _strConSql);

            _clsFun.FileLogDivFor("Divulvazione articoli", strFil, " *** DIVULGAZIONE ANAGRAFICA *** ");

            lblDiv.Text = "Controllo variazioni";
            progressBar1.Value = 0;
            progressBar1.Maximum = tabDiv.Rows.Count;
            progressBar1.Minimum = 0;

            foreach (DataRow y in tabDiv.Rows)
            {
                progressBar1.Increment(1);
                System.Windows.Forms.Application.DoEvents();

                if ((string)y["div_num"] == strNum)
                {
                    _clsFun.TypeDefault(tabDiv, y);

                    sMsg = "";

                    b = true;

                    //Se div_arf non definito divulgazione interna da terminalino
                    //if (!_clsFun.Numerico((string)y["div_arf"]))
                    if (((string)y["div_arf"]).Trim() == "")
                        b = false;

                    if (b)
                    {
                        //20170729 Dovrebbe servire nei casi che nelle righe seguenti con stesso codice manca la descrizione
                        j = tTmp.Select("div_arf='" + (string)y["div_arf"] + "'");
                        if (j.Length > 0 && !DBNull.Value.Equals(j[0]["div_art"]) && (string)j[0]["div_art"] != "")
                        {
                            y["div_art"] = (string)j[0]["div_art"];
                            //y["div_ard"] = (string)j[0]["div_ard"];

                            y["div_ard"] = "";
                            if(!DBNull.Value.Equals(j[0]["div_ard"]))
                                y["div_ard"] = (string)j[0]["div_ard"];
                        }

                        if ((string)y["div_art"] == "0167327")
                            Console.WriteLine("aaaaaaaaaa");

                        string sTipCosAmmessi = "";
                        string[] a = _str014AggPVenditaSoloFattura.Split(',');
                        if (a.Length > 1)
                            sTipCosAmmessi = a[1].Trim();

                        j = tAcq.Select("lia_arf='" + y["div_arf"] + "'", "lia_dti DESC");
                        if (strFil == "DIVAPOFFICE")    //Seck 20200515 Se ApOffice cerco con il codice articolo
                            j = tAcq.Select("lia_art='" + (string)y["div_art"] + "'", "lia_dti DESC");
                        if (j.Length > 0)
                        {
                            //Seck 20200417 cerco un costo valido del listino altrimenti prendo il primo

                            Boolean bb = false;

                            for (int i = 0; i < j.Length; i++ )
                            {
                                y["div_art"] = (string)j[i]["lia_art"];
                                y["art_cos"] = (decimal)j[i]["lia_cos"];

                                if (!DBNull.Value.Equals(j[i]["lia_dtf"]) && (DateTime)j[i]["lia_dtf"] != _clsDef.DAYOUT)
                                {
                                    if (DateTime.Compare(DateTime.Today, (DateTime)j[i]["lia_dti"]) >= 0 && DateTime.Compare(DateTime.Today, (DateTime)j[i]["lia_dtf"]) <= 0)
                                    {
                                        bb = true;
                                        break;
                                    }
                                }
                                //else if ((string)j[i]["lia_tip"] == "L" || (string)j[i]["lia_tip"] == "M")
                                else if (sTipCosAmmessi == "" || sTipCosAmmessi.Contains((string)j[i]["lia_tip"]))
                                {
                                    bb = true;
                                    break;
                                }
                            }
                            if(!bb)
                            {
                                y["div_art"] = (string)j[0]["lia_art"];
                                y["art_cos"] = (decimal)j[0]["lia_cos"];
                            }

                            //In alcuni casi non viene divulgato il pxc ... lo prendo se c'è dai lia precedenti
                            if ((decimal)y["div_pxc"] == 0)
                            {
                                for (int ii = 0; ii < j.Length; ii++)
                                {
                                    if ((decimal)j[ii]["lia_pxc"] > 0)
                                    {
                                        y["div_pxc"] = (decimal)j[ii]["lia_pxc"];
                                        break;
                                    }
                                }
                            }

                            if ((Boolean)j[0]["lia_ann"] == false)
                            {
                                //y["div_art"] = (string)j[0]["lia_art"];
                                //y["art_cos"] = (decimal)j[0]["lia_cos"];


                                ////In alcuni casi non viene divulgato il pxc ... lo prendo se c'è dai lia precedenti
                                //if ((decimal)y["div_pxc"] == 0)
                                //{
                                //    for (int ii = 0; ii < j.Length; ii++)
                                //    {
                                //        if ((decimal)j[ii]["lia_pxc"] > 0)
                                //        {
                                //            y["div_pxc"] = (decimal)j[ii]["lia_pxc"];
                                //            break;
                                //        }
                                //    }
                                //}


                                if ((decimal)j[0]["art_sfr"] > 0)
                                    y["art_sfr"] = (decimal)j[0]["art_sfr"];
                                if ((decimal)j[0]["art_tar"] > 0)
                                    y["art_tar"] = (decimal)j[0]["art_tar"];

                                if (((string)y["div_iva"]).Trim() == "")
                                    y["div_iva"] = (string)j[0]["art_iva"];

                                if (((String)y["div_ard"]).Trim() == "")
                                    y["div_ard"] = (string)j[0]["art_des"];
                                if (strFil != "DIVAPOFFICE")
                                    y["div_sta"] = (string)j[0]["art_sta"];

                                if (((String)y["div_rep"]).Trim() == "" && ((string)j[0]["art_rep"]).ToString() != "")
                                    y["div_rep"] = (string)j[0]["art_rep"];

                                if (((String)y["for_ecr"]).Trim() == "" && ((string)j[0]["art_ec3"]).ToString() != "")
                                {
                                    y["div_ec1"] = j[0]["art_ec1"];
                                    y["div_ec2"] = j[0]["art_ec2"];
                                    y["div_ec3"] = j[0]["art_ec3"];
                                }

                                if (((string)y["div_arf"]).Trim() == "223487")
                                    Console.WriteLine("");

                                //if (_strForRif == strFor)
                                //{
                                if (((string)y["for_umi"]).Trim() == "" && ((string)j[0]["art_umi"]).Trim() != "" && strFil != "DIVAPOFFICE")
                                    y["div_umi"] = j[0]["art_umi"];
                                if (((string)y["for_tgr"]).Trim() == "" && ((string)j[0]["art_tgr"]).Trim() != "" && strFil != "DIVAPOFFICE")
                                    y["div_tgr"] = j[0]["art_tgr"];
                                //}


                            }
                        }
                        else
                        {
                            DataRow[] j2 = _dasGen.Tables[TABTMPEAN].Select("die_for=" + strFor + " AND die_arf='" + (string)y["div_arf"] + "'");

                            if (j2.Length > 0)
                            {
                                for (int n = 0; n < j2.Length; n++)
                                {
                                    s = ((string)j2[n]["die_ean"]);

                                    if ((string)j2[n]["die_for"] != sForRif && s.Length == 13 && s.Substring(0, 1) == "2")
                                        Console.WriteLine("Salto");
                                    else
                                    {
                                        j = tEan.Select("ean_ean='" + (string)j2[n]["die_ean"] + "'");
                                        if (j.Length > 0)
                                        {
                                            y["div_art"] = (string)j[0]["ean_art"];

                                            if (strFil == "DIVAPOFFICE")    ///Seck 20210428 forzatura
                                                y["div_art"] = y["div_arf"];

                                            j = tAcq.Select("lia_art='" + (string)y["div_art"] + "'", "lia_dti DESC");
                                            if (j.Length > 0)
                                            {
                                                y["art_cos"] = (decimal)j[0]["lia_cos"];
                                                //Seck 20190219 altre info ...

                                                //y["div_art"] = (string)j[0]["lia_art"];
                                                y["art_cos"] = (decimal)j[0]["lia_cos"];
                                                y["art_sfr"] = (decimal)j[0]["art_sfr"];

                                                if (((string)y["div_iva"]).Trim() == "")
                                                    y["div_iva"] = (string)j[0]["art_iva"];

                                                if (((String)y["div_ard"]).Trim() == "")
                                                    y["div_ard"] = (string)j[0]["art_des"];
                                                y["div_sta"] = (string)j[0]["art_sta"];

                                                if (((String)y["div_rep"]).Trim() == "" && ((string)j[0]["art_rep"]).ToString() != "")
                                                    y["div_rep"] = (string)j[0]["art_rep"];

                                                if (((String)y["for_ecr"]).Trim() == "" && ((string)j[0]["art_ec3"]).ToString() != "")
                                                {
                                                    y["div_ec1"] = j[0]["art_ec1"];
                                                    y["div_ec2"] = j[0]["art_ec2"];
                                                    y["div_ec3"] = j[0]["art_ec3"];
                                                }

                                                if (((string)y["div_arf"]).Trim() == "223487")
                                                    Console.WriteLine("");

                                                if (((string)y["for_umi"]).Trim() == "" && ((string)j[0]["art_umi"]).Trim() != "" && strFil != "DIVAPOFFICE")
                                                    y["div_umi"] = j[0]["art_umi"];
                                                if (((string)y["for_tgr"]).Trim() == "" && ((string)j[0]["art_tgr"]).Trim() != "" && strFil != "DIVAPOFFICE")
                                                    y["div_tgr"] = j[0]["art_tgr"];


                                            }
                                            break;
                                        }
                                    }
                                }
                            }
                        }
                       

                        if (((string)y["div_art"]).Trim() != "")
                        {
                            if (((string)y["div_art"]).Trim() == "0013733")
                                Console.WriteLine("zzz");

                            /**/
                            j = tLiv.Select("liv_lis='" + _clsDef.LISPOS + "' AND liv_art='" + y["div_art"] + "'", "liv_dti DESC");
                            if (j.Length > 0 && (decimal)j[0]["liv_prv"] > 0 && !(Boolean)j[0]["liv_ann"] )
                            {
                                y["liv_lis"] = (string)j[0]["liv_lis"];
                                y["art_prp"] = (decimal)j[0]["liv_prv"];
                            }
                            else
                            {
                                j = tLiv.Select("liv_lis='" + _clsDef.LISFOR + "' AND liv_art='" + y["div_art"] + "'", "liv_dti DESC");
                                if (j.Length > 0)
                                {
                                    y["liv_lis"] = (string)j[0]["liv_lis"];
                                    y["art_prp"] = (decimal)j[0]["liv_prv"];
                                }
                            }

                            j = tLiv.Select("liv_lis='" + _clsDef.LISPRO + "' AND liv_art='" + y["div_art"] + "'", "liv_dti DESC");
                            if(j.Length > 0)
                            {
                                for(int i = 0; i < j.Length; i++)
                                {
                                    if((DateTime)j[i]["liv_dti"] <= DateTime.Today && (DateTime)j[i]["liv_dtf"] >= DateTime.Today)
                                    {
                                        y["div_off"] = ((decimal)j[0]["liv_prv"]).ToString("##0.00") + " PRO " + ((DateTime)j[0]["liv_dti"]).ToString("dd/MM/yy") + "-" + ((DateTime)j[0]["liv_dtf"]).ToString("dd/MM/yy");
                                    }
                                }
                            }

                            j = tOff.Select("ofa_art='" + (string)y["div_art"] + "'");
                            if (j.Length > 0)
                            {
                                y["div_off"] = ((decimal)j[0]["ofa_val"]).ToString("##0.00") + " " + (string)j[0]["ofa_cod"] + " " + ((DateTime)j[0]["oft_dti"]).ToString("dd/MM/yy") + "-" + ((DateTime)j[0]["oft_dtf"]).ToString("dd/MM/yy");
                            }
                             /**/
                        }

                        if ((string)y["div_arf"] == "000000006" || (string)y["div_arf"] == "792075")
                            Console.WriteLine("aaaaaaaa");

                        if (((string)y["div_iva"]).Trim() == "")
                           y["div_iva"] = CtrlTab(tTab, "IVA", y, "for_iva", strFor);

                        if (((string)y["div_umi"]).Trim() == "")
                            y["div_umi"] = CtrlTab(tTab, "UMI", y, "for_umi", strFor);

                        if (((string)y["div_tgr"]).Trim() == "")
                            y["div_tgr"] = CtrlTab(tTab, "TGR", y, "for_tgr", strFor);

                        if (((string)y["div_rep"]).Trim() == "")
                            y["div_rep"] = CtrlTab(tTab, "REP", y, "for_rep", strFor);

                        s = (string)y["div_art"];

                        if (((string)y["div_ec3"]).Trim() == "")
                        {
                            s = CtrlTab(tTab, "EC3", y, "for_ecr", strFor);
                            if (s.Length > 8)
                            {
                                y["div_ec1"] = s.Substring(0, 3);
                                y["div_ec2"] = s.Substring(3, 3);
                                y["div_ec3"] = s.Substring(6, 3);
                            }
                        }

                        y["dif_cos"] = (decimal)y["div_cos"] - (decimal)y["art_cos"];

                        //decimal dSfr = (decimal)y["art_sfr"];

                        if ((string)y["div_art"] == "0011973")
                            Console.WriteLine("zzzz");

                        d = 10;
                        j = tIva.Select("tab_cod='" + y["div_iva"] + "'");
                        if (j.Length > 0)
                            d = Convert.ToDecimal(j[0]["tab_ali"]);
                        y["art_mav"] = _clsFun.Margine((decimal)y["art_prp"], (decimal)y["div_cos"], d, (decimal)y["art_sfr"], "P");
                        y["div_mav"] = _clsFun.Margine((decimal)y["div_prp"], (decimal)y["div_cos"], d, (decimal)y["art_sfr"], "P");

                        if (sMsg != "")
                            _clsFun.FileLogDivFor("Divulvazione articoli", strFil, sMsg);
                    }
                    else if ((string)y["div_art"] != "")
                    {
                        //j = tAcq.Select("art_cod='" + y["div_art"] + "'");
                        j = tAcq.Select("lia_art='" + y["div_art"] + "'");
                        if (j.Length > 0)
                        {
                            y["div_sta"] = (string)j[0]["art_sta"];
                            y["div_iva"] = (string)j[0]["art_iva"];
                            y["div_umi"] = (string)j[0]["art_umi"];
                            y["div_tgr"] = (string)j[0]["art_tgr"];
                            y["div_rep"] = (string)j[0]["art_rep"];
                            //y["div_ec1"] = s.Substring(0, 3);
                            //y["div_ec2"] = s.Substring(3, 3);
                            //y["div_ec3"] = s.Substring(6, 3);
                            y["art_sfr"] = (decimal)j[0]["art_sfr"];

                            j = tLiv.Select("liv_lis='" + _clsDef.LISPOS + "' AND liv_art='" + (string)y["div_art"] + "'", "liv_dti DESC");
                            if (j.Length > 0 && (decimal)j[0]["liv_prv"] > 0)
                            {
                                y["liv_lis"] = (string)j[0]["liv_lis"];
                                y["art_prp"] = (decimal)j[0]["liv_prv"];

                            }
                            else
                            {
                                j = tLiv.Select("liv_lis='" + _clsDef.LISFOR + "' AND liv_art='" + (string)y["div_art"] + "'", "liv_dti DESC");
                                if (j.Length > 0 && (decimal)j[0]["liv_prv"] > 0)
                                {
                                    y["liv_lis"] = (string)j[0]["liv_lis"];
                                    y["art_prp"] = (decimal)j[0]["liv_prv"];

                                }
                            }

                            s = "SELECT lia_for, lia_arf, lia_art, lia_cos, lia_dti, lia_dtf, lia_ann FROM GesLisAcquisto ";
                            s += "WHERE lia_art='" + y["div_art"] + "' ORDER BY lia_dti DESC";
                            DataTable ttAcq = _clsFun.FillTabSql(TABLISACQ, s, true, _strConSql);
                            if (ttAcq.Rows.Count > 0)
                            {
                                decimal dSfr = (decimal)y["art_sfr"];

                                d = 10;
                                j = tIva.Select("tab_cod='" + y["div_iva"] + "'");
                                if (j.Length > 0)
                                    d = Convert.ToDecimal(j[0]["tab_ali"]);

                                y["art_cos"] = (decimal)ttAcq.Rows[0]["lia_cos"];
                                y["art_mav"] = _clsFun.Margine((decimal)y["art_prp"], (decimal)y["art_cos"], d, dSfr, "P");
                            }
                        }
                    }
                }
            }
            return tTab;
        }

        private string CtrlTab(DataTable tabTab, string strTip, DataRow rowDiv, string strFld, string strFor)
        {
            DataRow[] j;
            string s = "";
            string sCod = "";
            string sCof = ((string)rowDiv[strFld]).Trim();
            s = "tab_for='" + strFor + "' AND tab_tip='" + strTip + "' AND tab_cof='" + sCof + "'";
            j = tabTab.Select(s);
            if (j.Length > 0)
            {
                sCod = (string)j[0]["tab_cod"];
            }
            else
            {
                if ((string)rowDiv["div_art"] == "")
                    Console.WriteLine("aaaa");
                //MessageBox.Show("aaaaa");

                DataRow x = tabTab.NewRow();
                x["tab_for"] = strFor;
                x["tab_tip"] = strTip;
                x["tab_cof"] = sCof;
                x["tab_cod"] = "";
                //x["TabArt"] = ((string)rowDiv["div_art"]).Trim() + " " + (string)rowDiv["div_ard"];
                x["TabArt"] = ((string)rowDiv["div_ard"]).Trim()  + " - " + ((string)rowDiv["div_art"]).Trim();
                x["TabMsg"] = (string)rowDiv[strFld] + " codice non trovato ";
                tabTab.Rows.Add(x);
            }
            return sCod.Trim();
        }

        private void DivAggiorna()
        {
            //string sForRif = _clsFun.ParGet(clsDefine.enuParametri.ParForDefault, _strConSql);
            //if (sForRif.Length > 5)
            //    sForRif = sForRif.Substring(0, 5);

            string sTipForDiv = "";     //20200423 Seck Parametro in tab_fil di TabForImport per fornitore ApOffice o fornitore reale

            string s = "SELECT * FROM AnaArticoli ORDER BY art_cod";
            DataTable tArt = _clsFun.FillTabSql(TABANAART, s, false, _strConSql);
            DataColumn[] keys = new DataColumn[1];
            keys[0] = tArt.Columns["art_cod"];
            tArt.PrimaryKey = keys;

            s = "SELECT * FROM AnaBarcode ORDER BY ean_ean";
            DataTable tEan = _clsFun.FillTabSql(TABANAEAN, s, false, _strConSql);
            keys = new DataColumn[2];
            keys[0] = tEan.Columns["ean_ean"];
            keys[1] = tEan.Columns["ean_art"];
            tEan.PrimaryKey = keys;

            s = "SELECT * FROM GesLisAcquisto ORDER BY lia_arf, lia_for, lia_dti DESC";
            DataTable tAcq = _clsFun.FillTabSql(TABLISACQ, s, false, _strConSql);
            keys = new DataColumn[5];
            keys[0] = tAcq.Columns["lia_tip"];
            keys[1] = tAcq.Columns["lia_arf"];
            keys[2] = tAcq.Columns["lia_for"];
            keys[3] = tAcq.Columns["lia_dti"];
            keys[4] = tAcq.Columns["lia_dtf"];
            tAcq.PrimaryKey = keys;

            s = "SELECT ";
            s += "lia_art, ";
            s += "MAX(lia_dti) AS Dti, ";
            s += "MAX(lia_prv) AS lia_prv ";
            s += "FROM GesLisAcquisto ";
            s += "WHERE (lia_for = '00001') ";
            s += "GROUP BY lia_art";
            DataTable tAcp = _clsFun.FillTabSql(TABLISACQ, s, false, _strConSql);
            keys = new DataColumn[1];
            keys[0] = tAcp.Columns["lia_art"];
            tAcp.PrimaryKey = keys;

            s = "SELECT * FROM GesLisVendita ORDER BY liv_art, liv_dti DESC";
            DataTable tLiv = _clsFun.FillTabSql(TABLISVEN, s, false, _strConSql);
            keys = new DataColumn[4];
            keys[0] = tLiv.Columns["liv_lis"];
            keys[1] = tLiv.Columns["liv_art"];
            keys[2] = tLiv.Columns["liv_dti"];
            keys[3] = tLiv.Columns["liv_dtf"];
            tLiv.PrimaryKey = keys;

            s = "SELECT * FROM GesOffTestate";
            DataTable tOft = _clsFun.FillTabSql(TABGESOFT, s, false, _strConSql);
            keys = new DataColumn[2];
            keys[0] = tOft.Columns["oft_yea"];
            keys[1] = tOft.Columns["oft_cod"];
            tOft.PrimaryKey = keys;

            s = "SELECT * FROM GesOffArticoli WHERE ofa_ann=0";
            DataTable tOfa = _clsFun.FillTabSql(TABGESOFA, s, false, _strConSql);
            keys = new DataColumn[3];
            keys[0] = tOfa.Columns["ofa_yea"];
            keys[1] = tOfa.Columns["ofa_cod"];
            keys[2] = tOfa.Columns["ofa_rig"];
            tOfa.PrimaryKey = keys;

            s = "SELECT * FROM AnaArtIngredienti";
            DataTable tIng = _clsFun.FillTabSql("AnaArtIngredienti", s, false, _strConSql);
            keys = new DataColumn[3];
            keys[0] = tIng.Columns["ing_tip"];
            keys[1] = tIng.Columns["ing_row"];
            keys[2] = tIng.Columns["ing_art"];
            tIng.PrimaryKey = keys;

            DataTable tDie = _dasGen.Tables[TABTMPEAN];
            keys = new DataColumn[4];
            keys[0] = tDie.Columns["die_for"];
            keys[1] = tDie.Columns["die_arf"];
            keys[2] = tDie.Columns["die_ean"];
            keys[3] = tDie.Columns["die_num"];
            tDie.PrimaryKey = keys;

            ArrayList aVarArtDiv = new ArrayList();

            foreach(DataRow k in ((DataTable)dgv1.DataSource).Rows)
            {
                progressBar1.Value = 0;
                progressBar1.Maximum = _dasGen.Tables[TABTMPFDR].Rows.Count;
                progressBar1.Minimum = 0;

                _clsFun.TypeDefault(((DataTable)dgv1.DataSource), k);

                if ((string)k["dit_cho"] == "E")
                {
                    if ((string)k["dit_tva"] == VARART)
                    {
                        if ((string)k["dit_tfo"] == "APOFFICE")
                        {
                            s = "SELECT * FROM TabForImport WHERE tab_des='APOFFICE'";
                            DataTable tTmp = _clsFun.FillTabSql("", s, false, _strConSql);
                            if(tTmp.Rows.Count > 0)
                            {
                                s = (string)tTmp.Rows[0]["tab_fil"];
                                string[] a = s.Split('|');
                                if (a.Length > 2)
                                    sTipForDiv = a[2];
                            }
                        }

                        foreach (DataRow y in _dasGen.Tables[TABTMPFDR].Rows)
                        {
                            progressBar1.Increment(1);
                            System.Windows.Forms.Application.DoEvents();

                            if (DBNull.Value.Equals(y["div_art"]))
                                y["div_art"] = "";

                            if (DBNull.Value.Equals(y["div_var"]))
                                y["div_var"] = "";

                            if (DBNull.Value.Equals(y["div_eti"]))
                                y["div_eti"] = "";

                            if ((string)y["div_arf"] == "019781402")
                                Console.WriteLine("zzzzzzzz");

                            if ((string)y["div_art"] == "0011299")
                                Console.WriteLine("xxxxxxx");

                            if ((string)y["div_num"] == (string)k["dit_num"] && ((string)y["div_ard"]).Trim() != "" && ((string)y["div_var"]).Trim() != "N" && (_clsFun.Numerico((string)y["div_art"],"0123456789") || ((string)y["div_art"]).Trim() == ""))
                            {
                                DivAggArticolo(tArt, tEan, tAcq, tAcp, tLiv, tDie, y, k, true, sTipForDiv);
                                if (((string)y["div_var"] == "A"  && (string)y["div_sta"] == "A") || (string)y["div_eti"] == "S")
                                {
                                    _clsVar.Variazioni((string)y["div_art"], "Forza", "Divulg. " + (string)k["dit_des"], _clsDef.VARALL);

                                    s = (string)y["div_art"];
                                    if (aVarArtDiv.IndexOf(s) < 0)
                                        aVarArtDiv.Add(s);

                                }
                                if (DBNull.Value.Equals(y["div_ing"]))
                                    y["div_ing"] = "";
                                if ((string)y["div_ing"] != "")
                                    DivAggIngredienti(tIng, (string)y["div_art"], (string)y["div_ing"]);
                            }
                        }
                    }

                    if ((string)k["dit_tva"] == VARFAT)
                    {
                        foreach (DataRow y in _dasGen.Tables[TABTMPFDR].Rows)
                        {
                            progressBar1.Increment(1);
                            System.Windows.Forms.Application.DoEvents();

                            if (DBNull.Value.Equals(y["div_art"]))
                                y["div_art"] = "";

                            if (DBNull.Value.Equals(y["div_var"]))
                                y["div_var"] = "";

                            if (DBNull.Value.Equals(y["div_eti"]))
                                y["div_eti"] = "";

                            if ((string)y["div_art"] == "0015804")
                                Console.WriteLine("zzzzzzzz");
                            
                            if ((string)y["div_num"] == (string)k["dit_num"] && ((string)y["div_ard"]).Trim() != "" && ((string)y["div_var"]).Trim() != "N")
                            {
                                DivAggArticolo(tArt, tEan, tAcq, tAcp, tLiv, tDie, y, k, true,sTipForDiv);
                                if ((string)y["div_eti"] == "S")
                                    _clsVar.Variazioni((string)y["div_art"], "Forza", "Fatt. " + (string)k["dit_ndo"] + " " + (string)k["dit_des"], _clsDef.VARETI);
                            }

                            if ((string)y["div_var"] == "A")
                            {
                                s = (string)y["div_art"];
                                if (aVarArtDiv.IndexOf(s) < 0)
                                    aVarArtDiv.Add(s);
                            }
                        }
                    }

                    if ((string)k["dit_tva"] == VAROFF)
                    {
                        foreach (DataRow y in _dasGen.Tables[TABTMPFDR].Rows)
                        {
                            progressBar1.Increment(1);
                            System.Windows.Forms.Application.DoEvents();

                            if (DBNull.Value.Equals(y["div_art"]))
                                y["div_art"] = "";

                            if (DBNull.Value.Equals(y["div_var"]))
                                y["div_var"] = "";

                            if (DBNull.Value.Equals(y["div_eti"]))
                                y["div_eti"] = "";

                            if ((string)y["div_num"] == (string)k["dit_num"])
                            {
                                //Se DPIU non inserisco articolo SE vuoto perchè con l'offerta i dati sono incompleti
                                if ((String)y["div_art"] == "" && (string)k["dit_tfo"] != "DPIU" && ((string)y["div_var"]).Trim() != "N")
                                    DivAggArticolo(tArt, tEan, tAcq, tAcp, tLiv, tDie, y, k, false, sTipForDiv); //iprezzi vengono inseriti dall'offerta
                                if ((String)y["div_art"] != "")
                                    DivAggOfferta(tArt, tOft, tOfa, tAcq, tLiv, y, k);
                            }
                        }
                    }

                }

                string sFil = (string)k["dit_pth"] + "\\" + (string)k["dit_fil"];
                if (File.Exists(sFil))
                {
                    s = Path.GetDirectoryName(sFil) + "\\Old\\" + Path.GetFileName(sFil) + "_" + DateTime.Now.ToString("yyyyMMddHHmm");
                    if (File.Exists(s))
                        File.Delete(s);
                    File.Move(sFil, s);
                }
                //else if ((string)k["dit_cho"] == "P")
                Salva(k);
            }

            string sPar016PathDivNegozi = _clsFun.ParGet(clsDefine.enuParametri.Par016PathDivNegozi, _strConSql);
            if (sPar016PathDivNegozi != "")
                _clsVar.DivNegArticoli(sPar016PathDivNegozi, aVarArtDiv, null, "", "");

            foreach(DataRow k in ((DataTable)dgv1.DataSource).Rows)
            {
                //if ((string)k["dit_fil"] == TABTMPDIV.ToUpper())
                if (((string)k["dit_cho"] == "E" || (string)k["dit_cho"] == "P" || (string)k["dit_cho"] == "C") && (string)k["dit_fil"] == TABTMPDIV.ToUpper())
                    TmpDivSta((string)k["dit_pth"]);
            }
        }

        private void Salva(DataRow rowDit)
        {
            string s = "";
            string p = "";
            string sNum = (string)rowDit["dit_num"];
            string sCho = (string)rowDit["dit_cho"];
            DataRow x;
            DataRow[] j;

            p = "GesDivForTestate";
            s = "SELECT * FROM GesDivForTestate WHERE dit_num='" + sNum + "'";
            DataTable t = _clsFun.FillTabSql(p, s, false, _strConSql);
            s = "";
            if (t.Rows.Count == 0)
                s = _clsFun.SqlInsertRow(p, t, rowDit);
            else
                s = "UPDATE GesDivForTestate SET dit_cho='" + sCho + "' WHERE dit_num='" + sNum + "'";
            if(s != "")
                _clsFun.SqlWrite(s, _strConSql);

            p = "GesDivForArticoli";
            s = "SELECT * FROM GesDivForArticoli WHERE div_num='" + sNum + "'";
            t = _clsFun.FillTabSql(p, s, false, _strConSql);

            ArrayList aWhe = new ArrayList();
            aWhe.Add("div_num");
            aWhe.Add("div_nri");
            ArrayList aExl = new ArrayList();

            progressBar1.Value = 0;
            progressBar1.Maximum = _dasGen.Tables[TABTMPFDR].Rows.Count;
            progressBar1.Minimum = 0;

            if (sCho == "P")
            {
                foreach (DataRow y in _dasGen.Tables[TABTMPFDR].Rows)
                {
                    progressBar1.Increment(1);
                    System.Windows.Forms.Application.DoEvents();

                    if ((string)y["div_num"] == (string)rowDit["dit_num"])
                    {
                        if (DBNull.Value.Equals(y["div_art"]))
                            y["div_art"] = "";

                        s = (string)y["div_art"];
                        Console.WriteLine("aaaaaaaaaaaaa");

                        j = t.Select("div_num='" + (string)y["div_num"] + "' AND div_nri='" + (string)y["div_nri"] + "'");
                        if (j.Length == 0)
                            s = _clsFun.SqlInsertRow(p, t, y);
                        else
                            s = _clsFun.SqlUpdRowIdx(p, t, j[0], y, aExl);
                        if (s != "")
                            _clsFun.SqlWrite(s, _strConSql);
                    }

                }

                p = "GesDivForBarcode";
                s = "SELECT * FROM GesDivForBarcode WHERE die_num='" + sNum + "'";
                t = _clsFun.FillTabSql(p, s, false, _strConSql);

                aWhe = new ArrayList();
                aWhe.Add("die_num");
                aWhe.Add("die_ean");
                aExl = new ArrayList();

                progressBar1.Value = 0;
                progressBar1.Maximum = _dasGen.Tables[TABTMPEAN].Rows.Count;
                progressBar1.Minimum = 0;

                foreach (DataRow y in _dasGen.Tables[TABTMPEAN].Rows)
                {
                    progressBar1.Increment(1);
                    System.Windows.Forms.Application.DoEvents();

                    if ((string)y["die_num"] == (string)rowDit["dit_num"])
                    {
                        //if ((string)y["die_num"] == "020" && (string)y["die_ean"] == "8710437002844")
                        //    Console.WriteLine("xxxx");

                        j = t.Select("die_num='" + (string)y["die_num"] + "' AND die_ean='" + (string)y["die_ean"] + "'");
                        if (j.Length == 0)
                            s = _clsFun.SqlInsertRow(p, t, y);
                        else
                            s = _clsFun.SqlUpdRow(p, t, j[0], y, aWhe, aExl);
                        
                        if (s != "")
                        {
                            _clsFun.SqlWrite(s, _strConSql);
                            if (j.Length == 0)
                                t.ImportRow(y);
                        }

                        j = t.Select("die_num='" + (string)y["die_num"] + "' AND die_ean='" + y["die_ean"] + "'");
                    }
                }
            }
        }

        private void TmpDivSta(string strTim)
        {
            string s = "UPDATE TmpDiv SET tmp_sta='N' WHERE tmp_tip='ADV' AND tmp_day <= " + strTim;
            _clsFun.SqlWrite(s, _strConSql);
        }

        private void DivAggArticolo(DataTable tabArt, DataTable tabEan, DataTable tabAcq, DataTable tabAcp, DataTable tabLiv, DataTable tabDie, DataRow rowDiv, DataRow rowDit, Boolean bolPrz, string strTipForDiv)
        {
            string strFor = (string)rowDit["dit_for"];
            string strDes = (string)rowDit["dit_des"];
            string strTva = (string)rowDit["dit_tva"];
            string strTfo = (string)rowDit["dit_tfo"];

            ArrayList aWhe = new ArrayList();
            ArrayList aExl = new ArrayList();

            string sPar19VarForzata = _clsFun.ParGet(clsDefine.enuParametri.Par019VarForzata, _strConSql);
            string sEcrDef = _clsFun.ParGet(clsDefine.enuParametri.ParEcrDefault, _strConSql);
            if (sEcrDef.Length != 11)
                sEcrDef = "";

            DataRow x; 
            DataRow[] j;

            string s = "";
            string sSta  = "";

            Boolean b = true;

            string sArt = "";
            if (!DBNull.Value.Equals(rowDiv["div_art"]))
                sArt = ((string)rowDiv["div_art"]).Trim();

            if ((string)rowDiv["div_art"] == "0000031")
                Console.WriteLine("zzzz");

            if (DBNull.Value.Equals(rowDiv["liv_lis"]))
                rowDiv["liv_lis"] = _clsDef.LISFOR;          //Seck 20190129

            if (((string)rowDiv["div_arf"]).Trim() == "")
                b = false;

            if ((string)rowDiv["div_arf"] == "001385201")
                Console.WriteLine("aaaaa");

            if (sArt == "0000018")
                Console.WriteLine("aaaaa");

            if (b)
            {
                if (sArt == "")
                {
                    j = tabDie.Select("die_for='" + strFor + "' AND die_arf='" + (string)rowDiv["div_arf"] + "'");
                    sArt = _clsQry.SeekArfEan(strFor, (string)rowDiv["div_arf"], j); 
                    if (sArt == "")
                        sArt = _clsFun.NewNum(_clsDef.COD04Z, clsDefine.enuNumeratori.NumAnaArticoli, 7, _strConSql);
                }
                rowDiv["div_art"] = sArt;

                j = tabArt.Select("art_cod='" + sArt + "'");

                x = tabArt.NewRow();
                if (j.Length > 0)
                {
                    foreach (DataColumn c in tabArt.Columns)
                        x[c] = j[0][c];
                }
                x["art_cod"] = sArt;

                if (DBNull.Value.Equals(x["art_des"]) || (string)x["art_des"] == "")
                    x["art_des"] = (string)rowDiv["div_ard"];

                if (DBNull.Value.Equals(x["art_deb"]) || (string)x["art_deb"] == "")
                {
                    if (!DBNull.Value.Equals(rowDiv["div_adb"]) && (string)rowDiv["div_adb"] != "")
                        x["art_deb"] = (string)rowDiv["div_adb"];
                }

                if(_str015AggDesArtFromFornitore == "S" && _strForRif == strFor)
                {
                    if (!DBNull.Value.Equals(rowDiv["div_ard"]) && (string)rowDiv["div_ard"] != "")
                        x["art_des"] = (string)rowDiv["div_ard"];
                    if (!DBNull.Value.Equals(rowDiv["div_adb"]) && (string)rowDiv["div_adb"] != "")
                        x["art_deb"] = (string)rowDiv["div_adb"];
                }

                if (DBNull.Value.Equals(x["art_sta"]) || (string)x["art_sta"] == "")
                    x["art_sta"] = "N";             //Non attivo

                if (DBNull.Value.Equals(rowDiv["div_sta"]))
                    rowDiv["div_sta"] = "";

                if ((string)rowDiv["div_sta"] == _clsDef.STAATT)
                    x["art_sta"] = _clsDef.STAATT;
                else if (strTfo == "APOFFICE")
                    x["art_sta"] = (string)rowDiv["div_sta"];

                if (DBNull.Value.Equals(x["art_iva"]))
                    x["art_iva"] = "";
                if (DBNull.Value.Equals(x["art_umi"]))
                    x["art_umi"] = "";
                if (DBNull.Value.Equals(x["art_tgr"]))
                    x["art_tgr"] = "";
                if (DBNull.Value.Equals(x["art_pxc"]))
                    x["art_pxc"] = 0;
                if (DBNull.Value.Equals(x["art_pne"]))
                    x["art_pne"] = 0;
                if (DBNull.Value.Equals(x["art_sfr"]))
                    x["art_sfr"] = 0;
                if (DBNull.Value.Equals(x["art_tar"]))
                    x["art_tar"] = 0;
                if (DBNull.Value.Equals(x["art_rep"]))
                    x["art_rep"] = "";
                if (DBNull.Value.Equals(x["art_ec3"]))
                    x["art_ec3"] = "";

                if (!DBNull.Value.Equals(rowDiv["art_sfr"]))
                    x["art_sfr"] = (decimal)rowDiv["art_sfr"];
                if (!DBNull.Value.Equals(rowDiv["art_tar"]))
                    x["art_tar"] = (decimal)rowDiv["art_tar"];

                if (sArt == "0053723" || sArt == "00004063")
                    Console.WriteLine("aaaaaaa");

                if (!DBNull.Value.Equals(rowDiv["div_mar"]) && ((string)rowDiv["div_mar"]).Trim() != "")
                {
                    if (!DBNull.Value.Equals(rowDiv["div_mad"]) && ((string)rowDiv["div_mad"]).Trim() != "")
                    {
                        x["art_mar"] = (string)rowDiv["div_mar"];
                        CtrlTabMarchio((string)rowDiv["div_mar"], (string)rowDiv["div_mad"]);
                    }
                }

                if (_strForRif == strFor || ((string)x["art_iva"]).Trim() == "" || ((string)x["art_iva"]).Trim() == "000")
                {
                    if (DBNull.Value.Equals(rowDiv["div_iva"]))
                        rowDiv["div_iva"] = "";
                    if (((string)rowDiv["div_iva"]).Trim() != "")
                        x["art_iva"] = (string)rowDiv["div_iva"];
                }

                if (_strForRif == strFor || ((string)x["art_umi"]).Trim() == "" || ((string)x["art_umi"]).Trim() == "PZ")
                {
                    if (DBNull.Value.Equals(rowDiv["div_umi"]))
                        rowDiv["div_umi"] = "";
                    if (((string)rowDiv["div_umi"]).Trim() != "")
                        x["art_umi"] = (string)rowDiv["div_umi"];
                }

                if (_strForRif == strFor || ((string)x["art_tgr"]).Trim() == "")
                {
                    if (DBNull.Value.Equals(rowDiv["div_tgr"]))
                        rowDiv["div_tgr"] = "";

                    if (((string)rowDiv["div_tgr"]).Trim() != "")
                        x["art_tgr"] = (string)rowDiv["div_tgr"];
                }

                if (_strForRif == strFor || Convert.ToDecimal(x["art_pxc"]) == 0)
                {
                    if (DBNull.Value.Equals(rowDiv["div_pxc"]))
                        rowDiv["div_pxc"] = 0;

                    if (Convert.ToDecimal(rowDiv["div_pxc"]) > 0)
                        x["art_pxc"] = (decimal)rowDiv["div_pxc"];
                }

                if (_strForRif == strFor || Convert.ToDecimal(x["art_pne"]) == 0)
                {
                    if (DBNull.Value.Equals(rowDiv["div_pne"]))
                        rowDiv["div_pne"] = 0;

                    if (Convert.ToDecimal(rowDiv["div_pne"]) > 0)
                        x["art_pne"] = (decimal)rowDiv["div_pne"];
                }
                //Seck 20170407 Aggiorno il reparto dalla divulgazione
                //if ((string)rowDiv["div_rep"] != "" && (string)rowDiv["div_rep"] != (string)x["art_rep"])

                if (_strForRif == strFor || ((string)x["art_rep"]).Trim() == "")
                {
                    if (DBNull.Value.Equals(rowDiv["div_rep"]))
                        rowDiv["div_rep"] = "";

                    if (_strForRif == strFor && (string)rowDiv["div_rep"] != "")
                        x["art_rep"] = ((string)rowDiv["div_rep"]);
                    else if ((string)rowDiv["div_rep"] != "" && ((string)x["art_rep"]).Trim() == "" && (string)rowDiv["div_rep"] != (string)x["art_rep"])
                        x["art_rep"] = ((string)rowDiv["div_rep"]);
                }

                if (_strForRif == strFor || ((string)x["art_ec3"]).Trim() == "" || (sEcrDef.Length > 3 && (string)x["art_ec3"] == sEcrDef.Substring(0,3)))                
                {
                    string[] a = sEcrDef.Split(',');

                    if (DBNull.Value.Equals(rowDiv["div_ec1"]))
                        rowDiv["div_ec1"] = "";
                    if (DBNull.Value.Equals(rowDiv["div_ec2"]))
                        rowDiv["div_ec2"] = "";
                    if (DBNull.Value.Equals(rowDiv["div_ec3"]))
                        rowDiv["div_ec3"] = "";

                    //20200323 Tengo ECR di ApOffice
                    //if (DBNull.Value.Equals(x["art_ec1"]) || (string)x["art_ec1"] == "" || (string)x["art_ec1"] == "000" || (string)rowDiv["div_ec1"] != "")
                    if (DBNull.Value.Equals(x["art_ec1"]) || (string)x["art_ec1"] == "" || (string)x["art_ec1"] == "000" || _strPar017EcrDivFornitori == "S") // || (string)rowDiv["div_ec1"] != "")
                    {
                        if (a.Length != 3 || ((string)rowDiv["div_ec1"] != a[0]))
                            x["art_ec1"] = ((string)rowDiv["div_ec1"]);
                    }
                    if (DBNull.Value.Equals(x["art_ec2"]) || (string)x["art_ec2"] == "" || (string)x["art_ec2"] == "000" || _strPar017EcrDivFornitori == "S") // || (string)rowDiv["div_ec2"] != "")
                    {
                        if (a.Length != 3 || ((string)rowDiv["div_ec1"] != a[1]))
                            x["art_ec2"] = ((string)rowDiv["div_ec2"]);
                    }
                    if (DBNull.Value.Equals(x["art_ec3"]) || (string)x["art_ec3"] == "" || (string)x["art_ec3"] == "000" || _strPar017EcrDivFornitori == "S") // || (string)rowDiv["div_ec3"] != "")
                    {
                        if (a.Length != 3 || ((string)rowDiv["div_ec1"] != a[2]))
                            x["art_ec3"] = ((string)rowDiv["div_ec3"]);
                    }

                    if(a.Length > 1 && (DBNull.Value.Equals(x["art_ec3"]) || (string)x["art_ec3"] == ""))
                    {
                        x["art_ec1"] = a[0];
                        x["art_ec2"] = a[1];
                        x["art_ec3"] = a[2];
                    }

                }

                if (_strForRif == strFor)
                {
                    if (DBNull.Value.Equals(rowDiv["div_reb"]))
                        rowDiv["div_reb"] = "";

                    if ((!DBNull.Value.Equals(rowDiv["div_plu"]) && ((string)rowDiv["div_plu"]).Trim() != "") || (string)rowDiv["div_reb"] != "")
                    {
                        x["art_plu"] = (string)rowDiv["div_plu"];
                        x["art_reb"] = (string)rowDiv["div_reb"];

                        if ((string)rowDiv["div_bil"] == "S")
                            x["art_bil"] = true;

                        if (((string)rowDiv["div_ori"]).Trim() != "")
                            x["art_ori"] = (string)rowDiv["div_ori"];

                        if (((string)rowDiv["div_cal"]).Trim() != "")
                            x["art_cal"] = (string)rowDiv["div_cal"];

                        if (((string)rowDiv["div_cat"]).Trim() != "")
                            x["art_cat"] = (string)rowDiv["div_cat"];

                        if (((string)rowDiv["div_gsc"]).Trim() != "")
                            x["art_gsc"] = (string)rowDiv["div_gsc"];

                        if (((string)rowDiv["div_bpz"]).Trim() == "S")
                            x["art_bpz"] = true;

                        //Unicom
                        if (((string)rowDiv["for_ori"]).Trim() != "")
                            x["art_ori"] = CodFromTab("TabOrigine", ((string)rowDiv["for_ori"]).Trim());

                        if (((string)rowDiv["for_cal"]).Trim() != "")
                            x["art_cal"] = CodFromTab("TabCalibro", ((string)rowDiv["for_cal"]).Trim());

                        if (((string)rowDiv["for_cat"]).Trim() != "")
                            x["art_cat"] = CodFromTab("TabCategoria", ((string)rowDiv["for_cat"]).Trim());
                    }
                }

                if (strTva == VARFAT)
                    x["art_sta"] = _clsDef.STAATT;

                sSta = (string)x["art_sta"];

                if (!DBNull.Value.Equals(rowDiv["div_cos"]) && (decimal)rowDiv["div_cos"] > 0)
                    x["art_cos"] = (decimal)rowDiv["div_cos"];

                //Seck 20180502 se arriva da offerta non aggiorno il prezzo di vendita
                if (bolPrz && !DBNull.Value.Equals(rowDiv["div_prp"]) && (decimal)rowDiv["div_prp"] > 0)
                {
                    DataRow[] jj = tabLiv.Select("liv_lis='" + _clsDef.LISPOS + "' AND liv_art='" + sArt + "'");
                    if(jj.Length == 0)
                        x["art_prv"] = (decimal)rowDiv["div_prp"];
                }

                if (j.Length == 0)
                {
                    x["art_dti"] = DateTime.Today;
                    x["art_dtm"] = DateTime.Today;
                    s = _clsFun.SqlInsertRow(TABANAART, tabArt, x);
                }
                else
                {
                    x["art_dtm"] = DateTime.Today;
                    aWhe = new ArrayList();
                    aWhe.Add("art_cod");
                    aExl = new ArrayList();
                    aExl.Add("art_dti");
                    aExl.Add("art_dtm");

                    s = _clsFun.SqlUpdRow(TABANAART, tabArt, j[0], x, aWhe, aExl);
                }
                if (s != "")
                {
                    _clsFun.SqlWrite(s, _strConSql);

                    if (sSta == _clsDef.STAATT || sSta == _clsDef.STACAN)
                            _clsVar.Variazioni(sArt, s, "Divulg. " + strDes, _clsDef.VARALL);
                    _clsFun.FileLog("Divulg. " + strDes, sArt, s);

                    string[] aLog = { 
                                "VARART",                       //  log_tip
                                (string)x["art_cod"],           //  log_art
                                "",                             //  log_ean
                                (string)x["art_des"],           //  log_des
                                (string)rowDit["dit_fil"],      //  log_fil
                                "",                             //  log_acq
                                "",                             //  log_prv
                                ""                              //  log_msg
                                    };
                    _clsQry.LogSql(aLog);

                    if (j.Length > 0)
                        j[0].Delete();
                    tabArt.Rows.Add(x);
                }
                if (sSta == _clsDef.STAATT && sPar19VarForzata == "S")
                    _clsVar.Variazioni(sArt, "Forza", "Divulg. forzata " + strDes, _clsDef.VARALL);

                if ((string)rowDiv["div_art"] == "0000017" || (string)rowDiv["div_arf"] == "000349601")
                    Console.WriteLine("aaaaaaaa");

                s = "";
                DataRow[] j2 = tabDie.Select("die_for='" + strFor + "' AND die_arf='" + (string)rowDiv["div_arf"] + "'");
                if (j2.Length > 0)
                {
                    aWhe = new ArrayList();
                    aWhe.Add("ean_ean");
                    aExl = new ArrayList();
                    aExl.Add("ean_dti");
                    aExl.Add("ean_dtm");

                    for (int i = 0; i < j2.Length; i++)
                    {
                        //if ((string)j2[i]["die_ean"] == "8016926951045")
                        //    Console.WriteLine("aaaaa");

                        if (DBNull.Value.Equals(j2[i]["die_bil"]))
                            j2[i]["die_bil"] = false;

                        //Salto i barcode a peso di fornitori NON primari 20170614 Seck
                        //if ((string)j2[i]["die_for"] != _strForRif && s.Length == 13 && s.Substring(0, 1) == "2")
                            Console.WriteLine("Salto");

                        s = ((string)j2[i]["die_ean"]).Trim();

                        //Salto i barcode a peso di fornitori NON primari 20170614 Seck

                        b = true;
                        if (strDes == "APSHOP")
                            b = true;
                        else if ((string)j2[i]["die_for"]  != "00089" && (string)j2[i]["die_for"] != _strForRif && s.Length == 13 && s.Substring(0, 1) == "2" && !(Boolean)j2[i]["die_bil"])
                            b = false ;   //Console.WriteLine("Salto se barcode a peso a meno che non sia attivo per bilancia");
                        if(b)
                        {
                            if (sArt == "0071955")
                                Console.WriteLine("aaaaaaaaaaaa");

                            x = tabEan.NewRow();
                            x["ean_art"] = sArt;
                            x["ean_ean"] = ((string)j2[i]["die_ean"]).Trim();
                            
                            x["ean_qta"] = 0;
                            if (!DBNull.Value.Equals(j2[i]["die_qta"]))
                                x["ean_qta"] = (decimal)j2[i]["die_qta"];

                            x["ean_prv"] = 0;
                            if (!DBNull.Value.Equals(j2[i]["die_prv"]))
                                x["ean_prv"] = (decimal)j2[i]["die_prv"];

                            x["ean_bil"] = false;
                            if (!DBNull.Value.Equals(j2[i]["die_bil"]))
                                x["ean_bil"] = (Boolean)j2[i]["die_bil"];

                            x["ean_ecp"] = false;
                            if (!DBNull.Value.Equals(j2[i]["die_ecp"]))
                                x["ean_ecp"] = (Boolean)j2[i]["die_ecp"];

                            x["ean_ann"] = false;
                            if (!DBNull.Value.Equals(j2[i]["die_ann"]))
                                x["ean_ann"] = (Boolean)j2[i]["die_ann"];

                            x["ean_dti"] = DateTime.Today;
                            x["ean_dtm"] = DateTime.Today;

                            s = "";
                            j = tabEan.Select("ean_ean='" + j2[i]["die_ean"] + "'");
                            if (j.Length == 0)
                                s = _clsFun.SqlInsertRow(TABANAEAN, tabEan, x);
                            else
                            {
                                //if ((Boolean)x["ean_bil"] || (Boolean)j[0]["ean_bil"])
                                //    x["ean_bil"] = true;

                                s = _clsFun.SqlUpdRow(TABANAEAN, tabEan, j[0], x, aWhe, aExl);
                            }
                            //s = "UPDATE AnaBarcode SET ean_art='" + sArt + "', ean_dtm=" + _clsFun.DaySql(DateTime.Today) + " WHERE ean_ean='" + (string)j2[i]["die_ean"] + "'";
                            if (s != "")
                            {
                                _clsFun.SqlWrite(s, _strConSql);
                                if (sSta == _clsDef.STAATT)
                                    _clsVar.Variazioni(sArt, s, strDes, _clsDef.VARALL);
                                _clsFun.FileLog(strDes, sArt, s);

                                string[] aLog = { 
                                "VAREAN",                       //  log_tip
                                (string)x["ean_art"],           //  log_art
                                (string)x["ean_ean"],           //  log_ean
                                "",                             //  log_des
                                (string)rowDit["dit_fil"],      //  log_fil
                                "",                             //  log_acq
                                "",                             //  log_prv
                                ""                              //  log_msg
                                    };
                                _clsQry.LogSql(aLog);

                                if (j.Length > 0)
                                    j[0].Delete();
                                tabEan.Rows.Add(x);
                                rowDiv["div_var"] = "A";
                            }
                        }
                    }
                }

                if ((string)rowDiv["div_art"] == "0014632" || (string)rowDiv["div_arf"] == "748568")
                    Console.WriteLine("aaaaaaaa");
                if (sArt == "0013597")
                    Console.WriteLine("aaaa");

                /***/
                s = "";
                if (bolPrz)
                {
                    string sFor = strFor;
                    if (strTfo == "APOFFICE" && strTipForDiv == "FORARF")       //Seck 20200423 Fornitore ApOffice oppure fornitore reale
                    {
                        sFor = (string)rowDiv["div_fo2"];
                        DivFor2Anag(sFor, (string)rowDiv["div_fod"]);
                    }

                    string sTip = "L";
                    if (strTva == VARFAT)
                        sTip = "F";
                    else if (strTva == VAROFF)
                        sTip = "O";

                    j = tabAcq.Select("lia_tip='" + sTip + "' AND lia_for='" + sFor + "' AND lia_arf='" + (string)rowDiv["div_arf"] + "'", "lia_dti DESC");
                    if ((decimal)rowDiv["div_cos"] > 0 || j.Length == 0)
                    {
                        x = tabAcq.NewRow();
                        x["lia_tip"] = sTip;
                        x["lia_art"] = sArt;
                        x["lia_for"] = sFor;
                        if (!DBNull.Value.Equals(rowDiv["div_for"]) && ((string)rowDiv["div_for"]).Trim() != "")
                            x["lia_for"] = rowDiv["div_for"];

                        x["lia_dti"] = (DateTime)rowDiv["div_dva"];

                        x["lia_dtf"] = _clsDef.DAYOUT;
                        if (!DBNull.Value.Equals(rowDiv["div_sif"]) && (DateTime)rowDiv["div_sif"] != _clsDef.DAYOUT)
                        {
                            x["lia_dtf"] = (DateTime)rowDiv["div_sif"];
                            x["lia_tip"] = "O";
                        }

                        x["lia_arf"] = (string)rowDiv["div_arf"];
                        x["lia_cos"] = (decimal)rowDiv["div_cos"];
                        x["lia_pxc"] = (decimal)rowDiv["div_pxc"];
                        x["lia_cxp"] = 1;
                        x["lia_day"] = DateTime.Today;
                        x["lia_prv"] = (decimal)rowDiv["div_prp"];
                        
                        if (DBNull.Value.Equals(rowDiv["div_fo2"]))
                            rowDiv["div_fo2"] = "";
                        x["lia_fo2"] = (string)rowDiv["div_fo2"];

                        if (DBNull.Value.Equals(rowDiv["div_af2"]))
                            rowDiv["div_af2"] = "";
                        x["lia_af2"] = (string)rowDiv["div_af2"];

                        if (DBNull.Value.Equals(rowDiv["div_stf"]))
                            rowDiv["div_stf"] = "";

                        x["lia_stf"] = (string)rowDiv["div_stf"];

                        if (sArt == "0014632")
                            Console.WriteLine("aaaaa");

                        if (j.Length == 0)
                            s = _clsFun.SqlInsertRow(TABLISACQ, tabAcq, x);
                        else
                        {
                            aWhe = new ArrayList();
                            aWhe.Add("lia_tip");
                            aWhe.Add("lia_for");
                            aWhe.Add("lia_arf");
                            aWhe.Add("lia_dti");
                            aWhe.Add("lia_dtf");
                            aExl = new ArrayList();
                            aExl.Add("lia_dti");
                            aExl.Add("lia_dtm");

                            s = "";
                            /* 
                             Obiettivo:
                             Inserire se il costo è diverso dall'ultimo costo in modo che sia in testa.
                             Deve essere in chiave data divulgazione
                            
                             * Devo prendere l'ultimo costo in data presente.
                             * Se costo diverso inserisco.
                             
                             */

                            //Vedo se c'è il prezzo in data di variazione
                            int ij = 0;
                            for (int ii = 0; ii < j.Length; ii++)
                            {
                                if ((DateTime)j[ii]["lia_dti"] == (DateTime)x["lia_dti"])
                                {
                                    ij = ii;
                                    break;
                                }
                            }

                            b = false;

                            if (sArt == "0014632")
                                Console.WriteLine("aaaaa");

                            //20201107 Seck se il costo divulgato è diverso dall'ultimo costo inserisco
                            //j2 = tabAcq.Select("lia_art='" + (string)rowDiv["div_art"] + "' AND lia_dti <= " + _clsFun.DayMdb((DateTime)x["lia_dti"]), "lia_dti DESC");
                            //j2 = tabAcq.Select("lia_tip='" + sTip + "' AND lia_for<>'" + sFor + "' AND lia_art='" + (string)rowDiv["div_art"] + "'", "lia_dti DESC");
                            j2 = tabAcq.Select("lia_tip='" + sTip + "' AND lia_for<>'" + sFor + "' AND lia_art='" + (string)rowDiv["div_art"] + "'", "lia_dti DESC");
                            if (j2.Length > 0)
                            {
                                if (DateTime.Compare((DateTime)j2[0]["lia_dti"], (DateTime)j[ij]["lia_dti"]) >= 0 || (decimal)j[ij]["lia_cos"] != (decimal)rowDiv["div_cos"])
                                    b = true;
                            }
                            if (!b) //20201109 da togliere ( tip == '' )
                            {
                                j2 = tabAcq.Select("lia_tip='' AND lia_for='" + sFor + "' AND lia_art='" + (string)rowDiv["div_art"] + "'", "lia_dti DESC");
                                if (j2.Length > 0)
                                {
                                    if (DateTime.Compare((DateTime)x["lia_dti"], (DateTime)j2[0]["lia_dti"]) >= 0 || (decimal)x["lia_cos"] != (decimal)j2[0]["lia_cos"])
                                        b = true;
                                }
                            }

                            if (b || ((decimal)j[ij]["lia_cos"] != (decimal)rowDiv["div_cos"] || (decimal)j[ij]["lia_prv"] != (decimal)rowDiv["div_prp"]))
                            {
                                if ((DateTime)x["lia_dti"] != (DateTime)j[ij]["lia_dti"])
                                    s = _clsFun.SqlInsertRow(TABLISACQ, tabAcq, x);
                                else
                                    s = _clsFun.SqlUpdRow(TABLISACQ, tabAcq, j[ij], x, aWhe, aExl);
                                b = true;
                            }

                        }
                        if (s != "")
                        {
                            _clsFun.SqlWrite(s, _strConSql);

                            if (j.Length > 0)
                            {
                                for (int i = 0; i < j.Length; i++)
                                {
                                    //j[i].Delete();
                                    if ((DateTime)j[i]["lia_dti"] == (DateTime)x["lia_dti"])
                                        j[i].Delete();
                                }
                            }
                            try
                            {
                                tabAcq.Rows.Add(x);

                                string[] aLog = { 
                                "VARacq",                               //  log_tip
                                (string)x["lia_art"],                   //  log_art
                                "",                                     //  log_ean
                                "",                                     //  log_des
                                (string)rowDit["dit_fil"],              //  log_fil
                                ((decimal)x["lia_cos"]).ToString(),     //  log_acq
                                "",                                     //  log_prv
                                ""                                      //  log_msg
                                };
                                _clsQry.LogSql(aLog);
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine(ex.Message);
                            }
                        }

                        if(strTfo == "APOFFICE")
                        {
                            j = tabAcq.Select("lia_for='" + sFor + "' AND lia_art='" + (string)rowDiv["div_art"] + "' AND lia_arf<>'" + (string)rowDiv["div_arf"] + "'");
                            if(j.Length > 0)
                            {
                                s = "UPDATE GesLisAcquisto SET lia_ann=1 WHERE lia_art='" + (string)rowDiv["div_art"] + "' AND lia_arf<>'" + (string)rowDiv["div_arf"] + "'";
                                _clsFun.SqlWrite(s, _strConSql);
                            }
                        }

                    }

                    s = "";
                }
            }

            if (b || (strFor == "" && (string)rowDiv["div_art"] != "" || (decimal)rowDiv["var_prp"] > 0) || ((decimal)rowDiv["art_prp"] != (decimal)rowDiv["div_prp"]))
            {
                if ((string)rowDiv["div_arf"] == "051621")
                    Console.WriteLine("zzzz");

                b = true;

                // Se parametro agg prv solo da fattura aggiorno non aggiorno se articolo attivo e ci sono prezzi di vendita 

                if (strTva != VARFAT && _str014AggPVenditaSoloFattura.Length > 0 && _str014AggPVenditaSoloFattura.Substring(0,1) == "S" && (string)rowDiv["div_sta"] == _clsDef.STAATT && (decimal)rowDiv["var_prp"] == 0)
                {
                    j = tabLiv.Select("liv_lis='" + _clsDef.LISFOR + "' AND liv_art='" + sArt + "'");
                    if(j.Length > 0)
                        b = false;
                }

                if (b)
                {
                    if(DBNull.Value.Equals(rowDiv["div_prp"]))
                        rowDiv["div_prp"] = 0;
                    if(DBNull.Value.Equals(rowDiv["var_prp"]))
                        rowDiv["var_prp"] = 0;

                    if ((decimal)rowDiv["div_prp"] > 0 || (decimal)rowDiv["var_prp"] > 0)
                    {
                        for (int i = 0; i <= 1; i++)
                        {
                            string sLis = _clsDef.LISFOR;
                            string sLva = "";
                            decimal d = (decimal)rowDiv["div_prp"];
                            if (i == 1)
                            {
                                sLis = _clsDef.LISPOS;

                                if (DBNull.Value.Equals(rowDiv["var_prp"]))
                                    rowDiv["var_prp"] = 0;
                                d = (decimal)rowDiv["var_prp"];
                                if (!DBNull.Value.Equals(rowDiv["div_lva"]) && (string)rowDiv["div_lva"] == i.ToString("000"))
                                    sLva = (string)rowDiv["div_lva"];
                                if(strTfo == "APOFFICE")
                                    sLva = i.ToString("000"); 
                            }

                            b = true;
                            if (i == 0 && (decimal)rowDiv["div_prp"] == 0)
                                b = false;
                            if (i == 1 && (decimal)rowDiv["var_prp"] == 0 && sLva == "")
                                b = false;
                            if (i == 0 && strFor != _strForRif)
                            {
                                //Non aggiorno se il fornitore non è quello di riferimento e c'è un costo del fornitore di riferimento
                                //Qui si crea un collo di bottiglia
                                j = tabAcq.Select("lia_for='" + _strForRif + "' AND lia_art='" + sArt + "'");
                                //j = tabAcp.Select("lia_art='" + sArt + "'");
                                if (j.Length > 0 && !DBNull.Value.Equals(j[0]["lia_prv"]) && (decimal)j[0]["lia_prv"] != 0)
                                    b = false;
                            }
                            if (b)
                            {
                                s = "";
                                j = tabLiv.Select("liv_lis='" + sLis + "' AND liv_art='" + sArt + "'", "liv_dti DESC");
                                //j = tabLiv.Select("liv_lis='" + sLis + "' AND liv_art='" + sArt + "' AND liv_dti=" + _clsFun.DayMdb((DateTime)rowDiv["div_dva"]));

                                x = tabLiv.NewRow();
                                x["liv_lis"] = sLis;
                                x["liv_art"] = sArt;

                                x["liv_prv"] = d;

                                x["liv_day"] = DateTime.Today;
                                x["liv_dti"] = (DateTime)rowDiv["div_dva"];
                                x["liv_dtf"] = _clsDef.DAYOUT;

                                x["liv_ann"] = 0;
                                if(sLva != "")
                                    x["liv_ann"] = 1;

                                if (strTfo == "APOFFICE" && sLis == _clsDef.LISPOS && j.Length > 0)
                                {
                                    x["liv_prv"] = (decimal)j[0]["liv_prv"];

                                    x["liv_day"] = DateTime.Today;
                                    x["liv_dti"] = DateTime.Today; //(DateTime)j[0]["div_dva"];
                                    x["liv_dtf"] = _clsDef.DAYOUT;

                                    //x["liv_ann"] = 0;
                                    //if (sLva != "")

                                    x["liv_ann"] = 1;

                                    //sLva = "001";

                                }

                                if (j.Length == 0)
                                {
                                    if (sLva != "")
                                        s = "";
                                    else
                                        s = _clsFun.SqlInsertRow(TABLISVEN, tabLiv, x);
                                }
                                else if (sLva != "")
                                    s = "UPDATE GesLisVendita SET liv_ann=1 WHERE liv_lis='" + sLis + "' AND liv_art='" + sArt + "' AND liv_dti=" + _clsFun.DaySql((DateTime)j[0]["liv_dti"]);
                                else
                                {
                                    aWhe = new ArrayList();
                                    aWhe.Add("liv_lis");
                                    aWhe.Add("liv_art");
                                    aWhe.Add("liv_dti");
                                    aWhe.Add("liv_dtf");
                                    aExl = new ArrayList();
                                    aExl.Add("liv_day");

                                    s = "";

                                    //Vedo se c'è il prezzo in data di variazione
                                    int ij = 0;
                                    for (int ii = 0; ii < j.Length; ii++)
                                    {
                                        if ((DateTime)j[ii]["liv_dti"] == (DateTime)x["liv_dti"])
                                        {
                                            ij = ii;
                                            break;
                                        }
                                    }

                                    if ((decimal)j[ij]["liv_prv"] != d)
                                    {
                                        if ((DateTime)x["liv_dti"] != (DateTime)j[ij]["liv_dti"])
                                        {
                                            if ((DateTime)x["liv_dti"] > (DateTime)j[0]["liv_dti"])
                                            {
                                                string sCloseOld = "UPDATE GesLisVendita SET liv_dtf=" + _clsFun.DaySql(((DateTime)x["liv_dti"]).AddDays(-1)) + " WHERE liv_lis='" + sLis + "' AND liv_art='" + sArt + "' AND liv_dti=" + _clsFun.DaySql((DateTime)j[0]["liv_dti"]);
                                                _clsFun.SqlWrite(sCloseOld, _strConSql);
                                            }
                                            s = _clsFun.SqlInsertRow(TABLISVEN, tabLiv, x);
                                        }
                                        else
                                            s = _clsFun.SqlUpdRow(TABLISVEN, tabLiv, j[ij], x, aWhe, aExl);
                                        //s = "";
                                    }
                                }
                                if (s != "")
                                {
                                    _clsFun.SqlWrite(s, _strConSql);
                                    if (sSta == _clsDef.STAATT)
                                    {
                                        if ((string)rowDiv["liv_lis"] == sLis)
                                            _clsVar.Variazioni(sArt, s, strDes, _clsDef.VARALL);
                                        else if (sLis == _clsDef.LISPOS)     // && (decimal)j[0]["liv_prv"] != d
                                            _clsVar.Variazioni(sArt, s, strDes, _clsDef.VARALL);
                                    }
                                    _clsFun.FileLog(strDes, sArt, s);

                                    string[] aLog = { 
                                    "VARACQ",                               //  log_tip
                                    (string)x["liv_art"],                   //  log_art
                                    "",                                     //  log_ean
                                    "",                                     //  log_des
                                    (string)rowDit["dit_fil"],              //  log_fil
                                    "",                                     //  log_acq
                                    ((decimal)x["liv_prv"]).ToString(),     //  log_prv
                                    ""                                      //  log_msg
                                    };
                                    _clsQry.LogSql(aLog);

                                    if (j.Length > 0)
                                    {
                                        for (int i2 = 0; i2 < j.Length; i2++)
                                        {
                                            if((DateTime)j[i2]["liv_dti"] == (DateTime)x["liv_dti"])
                                                j[i2].Delete();
                                        }
                                    }
                                    tabLiv.Rows.Add(x);
                                    rowDiv["div_var"] = "A";
                                }
                            }
                        }
                    }
                }
            }
        }

        private void DivAggIngredienti(DataTable tabIng, string strArt, string strIng)
        {
            DataRow[] j;
            DataRow x;
            string sTip = "DVB";
            string s = "";

            ArrayList aWhe = new ArrayList();
            aWhe.Add("ing_tip");
            aWhe.Add("ing_art");
            aWhe.Add("ing_row");

            ArrayList aExl = new ArrayList();

            DataTable tTmp = tabIng.Clone();

            string[] a = strIng.Split('|');

            if(a.Length > 0)
            {
                int i = 0;

                foreach(string ss in a)
                {
                    if (ss.Length > 2)
                    {

                        string sCrt = ss.Substring(0, 1);
                        string sIng = ss.Substring(1).Trim();

                        if (sIng.Length > 0)
                        {
                            i++;

                            x = tTmp.NewRow();
                            x["ing_tip"] = sTip;
                            x["ing_row"] = i.ToString("000");
                            x["ing_crt"] = sCrt;
                            x["ing_art"] = strArt;
                            x["ing_txt"] = sIng;
                            x["ing_ann"] = false;
                            tTmp.Rows.Add(x);
                        }
                    }
                }

                j = tabIng.Select("ing_tip='" + sTip + "' AND ing_art='" + strArt + "'");

                if(j.Length > 0)
                {

                    for (int ii = 0; ii < j.Length; ii++)
                    {
                        DataRow[] jj = tTmp.Select("ing_tip='" + sTip + "' AND ing_row='" + (ii+1).ToString("000") + "' AND ing_art='" + strArt + "'");
                        if(jj.Length == 0)
                        {
                            x = tTmp.NewRow();
                            x["ing_tip"] = sTip;
                            x["ing_row"] = j[ii]["ing_row"];
                            x["ing_crt"] = j[ii]["ing_crt"];
                            x["ing_art"] = strArt;
                            x["ing_txt"] = j[ii]["ing_txt"];
                            x["ing_ann"] = true;
                            tTmp.Rows.Add(x);
                        }
                    }
                }

                foreach(DataRow y in tTmp.Rows)
                {
                    s = "";
                    j = tabIng.Select("ing_tip='" + sTip + "' AND ing_art='" + strArt + "'");
                    if (j.Length == 0)
                        s = _clsFun.SqlInsertRow("AnaArtIngredienti", tabIng, y);
                    else
                        s = _clsFun.SqlUpdRow("AnaArtIngredienti", tabIng, j[0], y, aWhe, aExl);

                    if (s != "")
                        _clsFun.SqlWrite(s, _strConSql);

                }

            }
        }

        private void DivAggOfferta(DataTable tabArt, DataTable tabOft, DataTable tabOfa, DataTable tabAcq, DataTable tabLiv, DataRow rowDiv, DataRow rowDit)
        {
            string s = "";
            DataRow[] j;
            ArrayList aWhe = new ArrayList();
            ArrayList aExl = new ArrayList();

            string sArt = ((string)rowDiv["div_art"]).Trim();
            string sCod = "";

            if ((string)rowDiv["div_art"] == "0000845")
                Console.WriteLine("aaaaaaa");

            DataRow x = tabOft.NewRow();
            x["oft_yea"] = ((DateTime)rowDit["dit_soi"]).Year.ToString();
            x["oft_cod"] = "";
            x["oft_num"] = rowDit["dit_ndo"];

            s = "Offerta " + ((string)rowDit["dit_des"]).Trim();
            if (((string)rowDit["dit_ndo"]).Trim() != "")
                s += " n. " + (string)rowDit["dit_ndo"];
            if (s.Length > 50)
                s = s.Substring(0,50);
            //x["oft_des"] = "Offerta " + ((string)rowDit["dit_des"]).Trim() + " n. " + (string)rowDit["dit_ndo"];
            x["oft_des"] = s;

            x["oft_dai"] = rowDit["dit_sii"];
            x["oft_daf"] = rowDit["dit_sif"];
            x["oft_dti"] = rowDit["dit_soi"];
            x["oft_dtf"] = rowDit["dit_sof"];
            x["oft_sta"] = _clsDef.STANOA;

            s = "oft_yea='" + x["oft_yea"] + "' AND ";
            s += "oft_num='" + x["oft_num"] + "'";
            s += " AND oft_dti=#" + ((DateTime)x["oft_dti"]).ToString("MM/dd/yyyy") + "# AND ";
            s += "oft_dtf=#" + ((DateTime)x["oft_dtf"]).ToString("MM/dd/yyyy") + "#";

            j = tabOft.Select(s);
            if (j.Length == 0)
            {
                sCod  = _clsFun.NewNum((string)x["oft_yea"], clsDefine.enuNumeratori.NumGesOfferte, 3, _strConSql);
                x["oft_cod"] = sCod;
                s = _clsFun.SqlInsertRow(TABGESOFT, tabOft, x);
                tabOft.Rows.Add(x);
            }
            else
            {
                sCod = (string)j[0]["oft_cod"];
                aWhe = new ArrayList();
                aExl = new ArrayList();
                aWhe.Add("oft_yea");
                aWhe.Add("oft_cod");
                s = _clsFun.SqlUpdRow(TABGESOFT, tabOft, j[0], x, aWhe, aExl);
            }
            if (s != "")
            {
                _clsFun.SqlWrite(s, _strConSql);
                _clsVar.Variazioni(sArt, s, "Offerta " + ((string)rowDit["dit_des"]).Trim(), _clsDef.VARALL);
                _clsFun.FileLog("Offerta " + ((string)rowDit["dit_des"]).Trim(), sArt, s);

            }

            string sYea = ((DateTime)rowDit["dit_soi"]).Year.ToString();
            string sRig = Convert.ToInt32(rowDiv["div_nri"]).ToString("0000");

            //Controllo se articolo già presente altrimenti aggiungo sull'ultima riga
            j = tabOfa.Select("ofa_yea='" + sYea + "' AND ofa_cod='" + sCod + "' AND ofa_art='" + sArt + "'");
            if(j.Length > 0)
                sRig = (string)j[0]["ofa_rig"];
            else
            {
                j = tabOfa.Select("ofa_yea='" + sYea + "' AND ofa_cod='" + sCod + "'", "ofa_rig DESC");
                if (j.Length == 0)
                    sRig = "0001";
                else
                    sRig = (Convert.ToInt16(j[0]["ofa_rig"]) + 1).ToString("0000");
            }

            x = tabOfa.NewRow();
            x["ofa_yea"] = sYea;
            x["ofa_cod"] = sCod;
            x["ofa_rig"] = sRig;
            x["ofa_art"] = rowDiv["div_art"];

            x["ofa_tip"] = _clsDef.OFAPRZ ;
            if (((string)rowDiv["off_tip"]).Trim() != "")
                x["ofa_tip"] = rowDiv["off_tip"];

            x["ofa_cos"] = rowDiv["div_cos"];
            if((decimal)x["ofa_cos"] == 0)
                x["ofa_cos"] = rowDiv["art_cos"];

            if((string)x["ofa_tip"] == _clsDef.OFASCO)
            {
                x["ofa_prv"] = rowDiv["div_prp"];
                x["ofa_val"] = rowDiv["div_imp"];
            }
            else
            {
                x["ofa_val"] = rowDiv["div_prp"];
            }

            x["ofa_xem"] = 0;
            x["ofa_xen"] = 0;
            x["ofa_ann"] = 0;
            if (!DBNull.Value.Equals(rowDiv["off_ann"]))
                x["ofa_ann"] = rowDiv["off_ann"];
            x["ofa_mix"] = 0;

            //j = tabOfa.Select("ofa_yea='" + x["ofa_yea"] + "' AND ofa_cod='" + x["ofa_cod"] + "' AND ofa_art='" + x["ofa_art"] + "'");
            j = tabOfa.Select("ofa_yea='" + sYea + "' AND ofa_cod='" + sCod + "' AND ofa_art='" + sArt + "'");
            if (j.Length == 0)
            {
                s = _clsFun.SqlInsertRow(TABGESOFA, tabOfa, x);
            }
            else
            {
                aWhe = new ArrayList();
                aExl = new ArrayList();
                aWhe.Add("ofa_yea");
                aWhe.Add("ofa_cod");
                aWhe.Add("ofa_art");
                s = _clsFun.SqlUpdRow(TABGESOFA, tabOfa, j[0], x, aWhe, aExl);
            }
            if (s != "")
            {
                _clsFun.SqlWrite(s, _strConSql);
                _clsVar.Variazioni(sArt, s, "Offerta " + ((string)rowDit["dit_des"]).Trim(), _clsDef.VARALL);
                _clsFun.FileLog("Offerta " + ((string)rowDit["dit_des"]).Trim(), sArt, s);

                if (j.Length > 0)
                    j[0].Delete();
                tabOfa.Rows.Add(x);


                string[] aLog = { 
                            "VAROFF",                               //  log_tip
                            (string)x["ofa_art"],                   //  log_art
                            "",                                     //  log_ean
                            "",                                     //  log_des
                            (string)rowDit["dit_fil"],              //  log_fil
                            "",                                     //  log_acq
                            ((decimal)x["ofa_val"]).ToString(),     //  log_prv
                            ""                                      //  log_msg
                                };
                _clsQry.LogSql(aLog);
            }

            if ((decimal)rowDiv["div_cos"] > 0)
            {
                x = tabAcq.NewRow();
                x["lia_tip"] = "O";
                x["lia_art"] = sArt;
                x["lia_for"] = rowDit["dit_for"];
                x["lia_dti"] = (DateTime)rowDit["dit_sii"];
                x["lia_dtf"] = (DateTime)rowDit["dit_sif"];
                x["lia_arf"] = (string)rowDiv["div_arf"];
                x["lia_cos"] = (decimal)rowDiv["div_cos"];
                x["lia_pxc"] = (decimal)rowDiv["div_pxc"];
                x["lia_cxp"] = 1;
                x["lia_day"] = DateTime.Today;
                x["lia_prv"] = (decimal)rowDiv["div_prp"];

                s = "lia_tip='O' AND ";
                s += "lia_for='" + rowDit["dit_for"] + "' AND ";
                s += "lia_arf='" + (string)rowDiv["div_arf"] + "' AND ";
                s += "lia_dti=" + _clsFun.DayMdb((DateTime)rowDit["dit_sii"]) + " AND ";
                s += "lia_dtf=" + _clsFun.DayMdb((DateTime)rowDit["dit_sif"]) + " ";

                j = tabAcq.Select(s);
                if (j.Length == 0)
                    s = _clsFun.SqlInsertRow(TABLISACQ, tabAcq, x);
                else
                {
                    aWhe = new ArrayList();
                    aWhe.Add("lia_tip");
                    aWhe.Add("lia_for");
                    aWhe.Add("lia_arf");
                    aWhe.Add("lia_dti");
                    aWhe.Add("lia_dtf");
                    aExl = new ArrayList();
                    aExl.Add("lia_dti");
                    aExl.Add("lia_dtm");

                    s = _clsFun.SqlUpdRow(TABLISACQ, tabAcq, j[0], x, aWhe, aExl);
                }
                if (s != "")
                {
                    _clsFun.SqlWrite(s, _strConSql);
                    _clsVar.Variazioni(sArt, s, (string)rowDit["dit_des"], _clsDef.VARALL);
                    _clsFun.FileLog((string)rowDit["dit_des"], sArt, s);

                    if (j.Length > 0)
                        j[0].Delete();
                    tabAcq.Rows.Add(x);
                }
            }

            /*** 
             * 
             * Seck 20170330 in anagrafica pesco direttamente dalla gestione offerte 
             * 
             * ***/
            //if ((decimal)rowDiv["div_prp"] > 0)
            //{
            //    s = "liv_lis='" + _clsDef.LISOFF + "' AND ";
            //    s += "liv_art='" + sArt + "' AND ";
            //    s += "liv_dti=" + _clsFun.DayMdb((DateTime)rowDit["dit_soi"]) + " AND ";
            //    s += "liv_dtf=" + _clsFun.DayMdb((DateTime)rowDit["dit_sof"]) + " ";

            //    j = tabLiv.Select(s);
            //    x = tabLiv.NewRow();
            //    x["liv_lis"] = _clsDef.LISOFF;
            //    x["liv_art"] = sArt;
            //    x["liv_prv"] = (decimal)rowDiv["div_prp"];
            //    x["liv_dti"] = (DateTime)rowDit["dit_soi"];
            //    x["liv_dtf"] = (DateTime)rowDit["dit_sof"];
            //    x["liv_ann"] = 0;
            //    s = "";
            //    if (j.Length == 0)
            //        s = _clsFun.SqlInsertRow(TABLISVEN, tabLiv, x);
            //    else
            //    {
            //        aWhe = new ArrayList();
            //        aWhe.Add("liv_lis");
            //        aWhe.Add("liv_art");
            //        aWhe.Add("liv_dti");
            //        aWhe.Add("liv_dtf");
            //        aExl = new ArrayList();

            //        if ((decimal)rowDiv["div_prp"] != (decimal)j[0]["liv_prv"])
            //        {
            //            if ((DateTime)x["liv_dti"] != DateTime.Today)
            //                s = _clsFun.SqlInsertRow(TABLISVEN, tabLiv, x);
            //            else
            //                s = _clsFun.SqlUpdRow(TABLISVEN, tabLiv, j[0], x, aWhe, aExl);
            //        }
            //    }
            //    if (s != "")
            //    {
            //        _clsFun.SqlWrite(s, _strConSql);
            //        //_clsVar.Variazioni(sArt, s, (string)rowDit["dit_des"], _clsDef.VARALL);
            //        _clsFun.FileLog((string)rowDit["dit_des"], sArt, s);

            //        if (j.Length > 0)
            //            j[0].Delete();
            //        tabLiv.Rows.Add(x);
            //    }
            //}
        }

        private void dgv1_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            if (dgv1.Columns[e.ColumnIndex].Name == "S")
            {
                string s = (string)dgv1.Rows[e.RowIndex].Cells[e.ColumnIndex].Value;
                if (s == "P")
                    e.CellStyle.BackColor = Color.Gold;
                else if (s == "C")
                    e.CellStyle.BackColor = Color.Red;
            }
        }

        private void logToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if(File.Exists(_clsDef.FILLOGDIVFOR))
            { 
                Process.Start("notepad.exe", _clsDef.FILLOGDIVFOR);
                if (MessageBox.Show("Cancello il LOG?", "", MessageBoxButtons.YesNo, MessageBoxIcon.Question) != DialogResult.Yes)
                {
                    File.Delete(_clsDef.FILLOGDIVFOR);
                }
            }
        }

        private void dgv2_EditingControlShowing(object sender, DataGridViewEditingControlShowingEventArgs e)
        {
            e.CellStyle.BackColor = Color.Aquamarine;

            if (e.Control is TextBox)
            {
                var txt = e.Control as TextBox;
                if (txt != null)
                {
                    e.Control.KeyPress += new KeyPressEventHandler(txt_KeyPress);
                }
            }
        }

        private void txt_KeyPress(object sender, KeyPressEventArgs e)
        {
            if ((sender as TextBox).Text != null && (sender as TextBox).Text.Trim() != "")
            {
                if (e.KeyChar == '.')
                    e.KeyChar = ',';
            }
        }

        private void dgv2_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            CurrencyManager cm = dgv2.BindingContext[dgv2.DataSource, dgv2.DataMember] as CurrencyManager;
            if (cm.Position >= 0)
            {
                DataRowView r = cm.List[dgv2.CurrentRow.Index] as DataRowView;
                DataRow x = r.Row;

                decimal dSfr = (decimal)x["art_sfr"];

                decimal d = 10;
                DataRow[] j = _dasGen.Tables[TABTABIVA].Select("tab_cod='" + x["div_iva"] + "'");
                if (j.Length > 0)
                    d = Convert.ToDecimal(j[0]["tab_ali"]);

                x["var_mav"] = _clsFun.Margine((decimal)x["var_prp"], (decimal)x["div_cos"], d, dSfr, "P");
            }
        }

        private void dgv2_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            if (dgv2.Columns[e.ColumnIndex].Name == "Variazione")
            {
                if (!DBNull.Value.Equals(dgv2.Rows[e.RowIndex].Cells[e.ColumnIndex].Value))
                {
                    string s = (string)dgv2.Rows[e.RowIndex].Cells[e.ColumnIndex].Value;
                    if (s == "N")
                        e.CellStyle.BackColor = Color.Red;
                }
            }
            else if (dgv2.Columns[e.ColumnIndex].Name == "Differenza costi")
            {
                //decimal d2 = (decimal)dgv2.Rows[e.RowIndex].Cells[e.ColumnIndex-2].Value;

                if ( !DBNull.Value.Equals( dgv2.Rows[e.RowIndex].Cells[e.ColumnIndex - 2].Value))
                {
                    decimal d2 = (decimal)dgv2.Rows[e.RowIndex].Cells[e.ColumnIndex - 2].Value;

                    decimal d = Convert.ToDecimal(e.Value);

                    if (e.Value != null && Math.Abs(d) > Convert.ToDecimal(0.001))
                    {
                        if (d2 == 0)
                            e.CellStyle.BackColor = Color.PeachPuff;
                        else if (Convert.ToDecimal(e.Value) > 0)
                            e.CellStyle.BackColor = Color.LightCoral;
                        else if (Convert.ToDecimal(e.Value) < 0)
                            e.CellStyle.BackColor = Color.LightGreen;
                    }
                }
            }
            else if (dgv2.Columns[e.ColumnIndex].Name == "Costo")
                e.CellStyle.BackColor = Color.Gold;
            else if (dgv2.Columns[e.ColumnIndex].Name == "Nuovo Prezzo")
                e.CellStyle.BackColor = Color.GreenYellow;
            else if (dgv2.Columns[e.ColumnIndex].Name == "Prezzo anagrafico")
                e.CellStyle.BackColor = Color.PowderBlue;
            else if (dgv2.Columns[e.ColumnIndex].Name == "Prezzo consigliato")
                e.CellStyle.BackColor = Color.PowderBlue; 
        }

        private void dgv2_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            if (dgv2.Columns[e.ColumnIndex].Name == "Descrizione")
            {
                CurrencyManager cm = dgv2.BindingContext[dgv2.DataSource, dgv2.DataMember] as CurrencyManager;
                if (cm.Position >= 0)
                {
                    DataRowView r = cm.List[dgv2.CurrentRow.Index] as DataRowView;
                    DataRow x = r.Row;

                    string s = "";
                    if (!DBNull.Value.Equals(x["div_art"]))
                        s = ((string)x["div_art"]).Trim();

                    if (s != "")
                    {
                        frmAnaArticolo f = new frmAnaArticolo();
                        f._strArtCod = s;
                        f.ShowDialog();

                        if(f._strSqlArt != "" && f._tabArt.Rows.Count > 0)
                        {
                            DataRow[] j;
                            DataTable tDiv = ((DataView)dgv2.DataSource).Table;

                            //if (f._strSqlArt.Contains("art_rep") || f._strSqlArt.Contains("art_umi") || f._strSqlArt.Contains("art_ec1"))
                            if (f._strSqlArt != "")
                            {
                                DataTable t = f._tabArt;
                                j = tDiv.Select("div_art='" + (string)t.Rows[0]["art_cod"] + "'");
                                if (j.Length > 0)
                                {
                                    for (int i = 0; i < j.Length; i++)
                                    {
                                        if (f._strSqlArt.Contains("art_rep"))
                                            j[i]["div_rep"] = (string)t.Rows[0]["art_rep"];

                                        if (f._strSqlArt.Contains("art_umi"))
                                            j[i]["div_umi"] = (string)t.Rows[0]["art_umi"];

                                        if (f._strSqlArt.Contains("art_tgr"))
                                            j[i]["div_tgr"] = (string)t.Rows[0]["art_tgr"];

                                        if (f._strSqlArt.Contains("art_pne"))
                                            j[i]["div_pne"] = (decimal)t.Rows[0]["art_pne"];

                                        if (f._strSqlArt.Contains("art_ec3"))
                                        {
                                            j[i]["div_ec1"] = (string)t.Rows[0]["art_ec1"];
                                            j[i]["div_ec2"] = (string)t.Rows[0]["art_ec2"];
                                            j[i]["div_ec3"] = (string)t.Rows[0]["art_ec3"];
                                        }
                                    }
                                }

                                x["div_var"] = "A";
                            }
                        }
                    }
                }
            }
            else if (dgv2.Columns[e.ColumnIndex].Name == "Varia")
            {

            }
        }

        private void dgv2_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.ColumnIndex == 1)
            {
                if (e.RowIndex >= 0)
                {
                    string s = dgv2.Rows[e.RowIndex].Cells[e.ColumnIndex].Value.ToString();

                    if (s == "")
                        dgv2.Rows[e.RowIndex].Cells[e.ColumnIndex].Value = "N";
                    else
                        dgv2.Rows[e.RowIndex].Cells[e.ColumnIndex].Value = "";
                }
            }


            else if (e.ColumnIndex == 18)
            {
                if (e.RowIndex >= 0)
                {
                    CurrencyManager cm = dgv2.BindingContext[dgv2.DataSource, dgv2.DataMember] as CurrencyManager;
                    if (cm.Position >= 0)
                    {
                        DataRowView r = cm.List[dgv2.CurrentRow.Index] as DataRowView;
                        DataRow x = r.Row;

                        string s = (string)x["div_eti"];
                        if (s == "")
                            x["div_eti"] = "S";
                        else if (s == "S")
                            x["div_eti"] = "";
                    }
                }
            }
        }

        private void dgv2_CurrentCellChanged(object sender, EventArgs e)
        {
            lblCos.Text = "";

            try
            {
                CurrencyManager cm = dgv1.BindingContext[dgv1.DataSource, dgv1.DataMember] as CurrencyManager;
                if (cm.Position >= 0)
                {

                    DataRowView r = cm.List[dgv1.CurrentRow.Index] as DataRowView;
                    DataRow x = r.Row;

                    string sFor = (string)x["dit_for"];



                    cm = dgv2.BindingContext[dgv2.DataSource, dgv2.DataMember] as CurrencyManager;
                    if (cm.Position >= 0)
                    {
                        r = cm.List[dgv2.CurrentRow.Index] as DataRowView;
                        x = r.Row;

                        if ((string)x["div_art"] == "0013325")
                            Console.WriteLine("xxxxxx");

                        string s = "SELECT * FROM GesLisAcquisto WHERE lia_for='" + sFor + "' AND lia_art='" + (string)x["div_art"] + "'  ORDER BY lia_dti DESC";
                        DataTable t = _clsFun.FillTabSql(TABLISACQ, s, false, _strConSql);

                        s = "";
                        int i = 0;
                        foreach(DataRow y in t.Rows)
                        {
                            i++;
                            if (i > 9)
                                break;
                            s += (string)y["lia_tip"] +" "+ ((DateTime)y["lia_dti"]).ToString("dd/MM/yy") +" "+ ((decimal)y["lia_cos"]).ToString("#0.00") + " - ";
                        }

                        s = s.Substring(0, s.Length - 3);

                        lblCos.Text = s;
                    }

                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }

        private void DivFatture()
        {
            lblDiv.Text = "Documenti di acquisto";

            foreach (DataRow k in ((DataTable)dgv1.DataSource).Rows)
            {
                if ((string)k["dit_cho"] == "E")
                {
                    if ((string)k["dit_tva"] == VARFAT || (string)k["dit_tva"] == VARNCA)
                        DivFatMovimenti(k);
                }
                //else if ((string)k["dit_cho"] == "P")
                Salva(k);
            }
        }

        private void DivFatMovimenti(DataRow rowDit)
        {
            DataRow[] j;
            ArrayList aWhe = new ArrayList();
            ArrayList aExl = new ArrayList();

            string sNum = (string)rowDit["dit_num"];
            string sFod = (string)rowDit["dit_des"];

            string sYfa = ((DateTime)rowDit["dit_ddo"]).Year.ToString();
            
            string sNeg = "001";
            if ((string)rowDit["dit_neg"] != "")
                sNeg = (string)rowDit["dit_neg"];

            string sUbi = _clsFun.FileIni("R", clsDefine.enuIni.Ini13Ubicazione, "");
            if (sUbi == "")
                sUbi = sNeg;

            string sNfa = "";
            string sFor = (string)rowDit["dit_for"];

            string sTpd = "FA";

            string sNdo = (string)rowDit["dit_ndo"];
            DateTime dDdo = (DateTime)rowDit["dit_ddo"];

            string s = "";
            s = "SELECT * FROM GesFatTestate WHERE ";
            s += "fat_ubi='" + sUbi + "' AND ";
            s += "fat_yfa='" + sYfa + "' AND ";
            s += "fat_cfo='" + sFor + "' AND ";
            s += "fat_tpd='" + sTpd + "' AND ";
            s += "fat_ndo='" + sNdo + "'";
            DataTable t = _clsFun.FillTabSql(TABGESFAT, s, false, _strConSql);

            DataRow y = t.NewRow();
            y["fat_ubi"] = sUbi;
            y["fat_yfa"] = sYfa;
            y["fat_nfa"] = sNfa;
            y["fat_neg"] = sNeg;
            y["fat_tpd"] = sTpd;

            y["fat_tdo"] = "FA";
            if ((string)rowDit["dit_tva"] == VARNCA)
                y["fat_tdo"] = "NA";

            y["fat_ndo"] = sNdo;
            y["fat_ddo"] = dDdo;
            y["fat_cfo"] = sFor;
            y["fat_no1"] = "";
            y["fat_tpg"] = "";
            y["fat_ann"] = 0;

            if (t.Rows.Count == 0)
            {
                sNfa = _clsFun.NewNum(DateTime.Today.Year.ToString("0000"), clsDefine.enuNumeratori.NumGesMovFatture, 6, _strConSql);
                y["fat_nfa"] = sNfa;
                s = _clsFun.SqlInsertRow(TABGESFAT, t, y);
            }
            else
            {
                sNfa = (string)t.Rows[0]["fat_nfa"];
                y["fat_nfa"] = sNfa;
                aWhe = new ArrayList();
                aWhe.Add("fat_ubi");
                aWhe.Add("fat_yfa");
                aWhe.Add("fat_nfa");
                aWhe.Add("fat_cfo");
                aExl = new ArrayList();
                s = _clsFun.SqlUpdRow(TABGESFAT, t, t.Rows[0], y, aWhe, aExl);
            }
            if (s != "")
            {
                _clsFun.SqlWrite(s, _strConSql);
                _clsFun.FileLog("GesFatTestate", "For: " + sFor + " Doc: " + sNdo, s);
            }

            s = "SELECT * FROM GesMovimenti WHERE ";
            s += "mov_ubi = '" + sUbi + "' AND ";
            s += "mov_yfa = '" + sYfa + "' AND ";
            s += "mov_nfa = '" + sNfa + "' "; 
            s += "ORDER BY mov_rfa";
            //t = _clsFun.FillTabSql(TABGESMOV, s, true, _strConSql);

            //DataView v = new DataView(_dasGen.Tables[TABTMPFDR], "div_tva='" + VARFAT + "' AND div_num='" + sNum + "'", "div_nri", DataViewRowState.CurrentRows);

            aWhe = new ArrayList();
            aWhe.Add("mov_ubi");
            aWhe.Add("mov_yfa");
            aWhe.Add("mov_nfa");
            aWhe.Add("mov_rfa");
            //aWhe.Add("mov_art");
            aExl = new ArrayList();

            t = _clsFun.FillTabSql(TABGESMOV, s, false, _strConSql);

            DataView v = new DataView(_dasGen.Tables[TABTMPFDR], "div_tva='" + (string)rowDit["dit_tva"] + "' AND div_num='" + sNum + "'", "div_nri", DataViewRowState.CurrentRows);

            foreach (DataRowView r in v)
            {
                if (DBNull.Value.Equals(r["div_qta"]))
                    r["div_qta"] = 0;

                if ((decimal)r["div_qta"] != 0)
                {
                    DataRow x = t.NewRow();
                    x["mov_ubi"] = sUbi;
                    x["mov_yfa"] = sYfa;
                    x["mov_nfa"] = sNfa;
                    x["mov_ymo"] = _clsDef.COD04X;
                    x["mov_nmo"] = _clsDef.COD06X;
                    x["mov_rfa"] = Convert.ToInt32(r["div_nri"]).ToString("0000");
                    x["mov_rmo"] = _clsDef.COD04X;
                    x["mov_art"] = r["div_art"];
                    x["mov_ard"] = r["div_ard"];
                    x["mov_iva"] = r["div_iva"];
                    x["mov_umi"] = r["div_umi"];
                    
                    if(sFod == "VENFRI")
                    {
                        x["mov_qta"] = 1;
                        x["mov_qkg"] = 0;

                        if ((string)r["div_umi"] == "KG")
                            x["mov_qkg"] = r["div_qta"];
                        else
                            x["mov_qta"] = r["div_qta"];

                    }
                    else
                        x["mov_qta"] = r["div_qta"];

                    x["mov_cos"] = r["div_cos"];
                    x["mov_prv"] = r["div_prp"];
                    x["mov_imp"] = r["div_imp"];
                    x["mov_ann"] = false;
                    x["mov_day"] = DateTime.Today;
                    x["mov_no1"] = "";

                    if (DBNull.Value.Equals(x["mov_imp"]) || (decimal)x["mov_imp"] == 0)
                        x["mov_imp"] = (decimal)x["mov_cos"] * (decimal)x["mov_qta"];

                    //s = "mov_art='" + x["mov_art"] + "'";
                    s = "mov_rfa='" + x["mov_rfa"] + "'";

                    j = t.Select(s);
                    if (j.Length > 0)
                    {
                        s = _clsFun.SqlUpdRow(TABGESMOV, t, j[0], x, aWhe, aExl);
                    }
                    else
                    {
                        s = _clsFun.SqlInsertRow(TABGESMOV, t, x);
                    }
                    if (s != "")
                    {
                        _clsFun.SqlWrite(s, _strConSql);
                        _clsFun.FileLog("GesMovimenti", (string)x["mov_art"], s);
                    }
                }
            }
        }

        private void Pulizia()
        {
            string s = "SELECT * FROM GesDivForTestate WHERE dit_cho='E' OR dit_cho='C'";
            DataTable t = _clsFun.FillTabSql("GesDivForTestate", s, false, _strConSql);
            foreach(DataRow y in t.Rows)
            {
                s = "DELETE GesDivForBarcode WHERE die_num='" + (string)y["dit_num"] + "'";
                _clsFun.SqlWrite(s, _strConSql);
                s = "DELETE GesDivForArticoli WHERE div_num='" + (string)y["dit_num"] + "'";
                _clsFun.SqlWrite(s, _strConSql);
                s = "DELETE GesDivForTestate WHERE dit_num='" + (string)y["dit_num"] + "'";
                _clsFun.SqlWrite(s, _strConSql);
            }
        }

        private void downloadToolStripMenuItem_Click(object sender, EventArgs e)
        {
            ForDownLoad();
        }

        public void ForDownLoad()
        {
            DataRow[] j;
            Boolean b = true;

            string s = "SELECT * FROM TabForImport WHERE tab_ann=0";
            DataTable t = _clsFun.FillTabSql("TabForImport", s, false, _strConSql);

            foreach (DataRow y in t.Rows)
            {
                //MessageBox.Show("'Ftp host' non definito in tabella import da fornitori!", "CONTROLLO FORNITORE", MessageBoxButtons.OK, MessageBoxIcon.Question);
                //else

                if (((string)y["tab_hst"]).Trim() != "")
                {

                    string sPathFor = ((string)y["tab_pth"]).Trim();

                    string[] sFils = Directory.GetFiles(sPathFor);
                    foreach (string sFil in sFils)
                    {
                        j = t.Select("tab_fil='" + Path.GetFileName(sFil).Substring(0, 7).ToUpper() + "'");
                        if (j.Length > 0)
                            Rinomina(sFil);
                        //b = false;

                        j = t.Select("tab_fat='" + Path.GetFileName(sFil).Substring(0, 7).ToUpper() + "'");
                        if (j.Length > 0)
                            Rinomina(sFil);
                        //b = false;

                        j = t.Select("tab_fio='" + Path.GetFileName(sFil).Substring(0, 7).ToUpper() + "'");
                        if (j.Length > 0)
                            Rinomina(sFil);
                        //b = false;
                    }

                    if (!b)
                        MessageBox.Show("Ci sono documenti non ancora elaborati di " + ((string)y["tab_des"]).Trim() + "!");
                    else
                    {
                        s = LeggiFtp(y);
                        if (s != "")
                            MessageBox.Show(s);
                    }
                }
            }
        }

        private void Rinomina(string strFil)
        {
            if (strFil.Length > 5 && strFil.Substring(strFil.Length - 3,1) != "_")
            {
                string sPath = Path.GetDirectoryName(strFil);
                string s = "";

                for (int i = 0; i < 100; i++)
                {
                    s = strFil + "_" + i.ToString("00");
                    if (!File.Exists(s))
                        break;
                }

                File.Move(strFil, s);
            }
        }

        private string LeggiFtp(DataRow rowDiv)
        {
            Color bc = lblDiv.BackColor;
            lblDiv.BackColor = Color.LightCoral;
            lblDiv.Text = "Ricezione in corso";
            System.Windows.Forms.Application.DoEvents();

            string s = "";
            string sMsg = "";
            Boolean b = false;
            string[] a = sMsg.Split('|');

            ArrayList aPre = new ArrayList();
            if (((string)rowDiv["tab_fil"]).Trim() != "")
            {
                //Nel caso di separatore "," distinguo il percorso
                //Nel caso di separatore "|" ci sono più prefissi di files

                s = ((string)rowDiv["tab_fil"]).Trim();

                a = s.Split(',');
                s = a[0];

                a = s.Split('|');
                foreach (string ss in a)
                {
                    if (ss.Trim() != "")
                        aPre.Add(ss);
                }
            }
            if (((string)rowDiv["tab_fat"]).Trim() != "")
                aPre.Add(((string)rowDiv["tab_fat"]).Trim());
            if (((string)rowDiv["tab_fio"]).Trim() != "")
                aPre.Add(((string)rowDiv["tab_fio"]).Trim());
            if(((string)rowDiv["tab_tip"]).Trim() == "FTPDIV")
                aPre.Add(((string)rowDiv["tab_ord"]).Trim());

            clsFtp clsFtp = new clsFtp();

            //clsFtp._strFtpHost = ((string)rowDiv["tab_hst"]).Trim();

            s = ((string)rowDiv["tab_hst"]).Trim();
            a = s.Split('|');
            string strHst = a[0];

            clsFtp._strFtpHost = strHst;
            clsFtp._strFtpUser = ((string)rowDiv["tab_usr"]).Trim();
            clsFtp._strFtpPswd = ((string)rowDiv["tab_pwd"]).Trim();
            //clsFtp._strFtpPath = ((string)rowDiv["tab_fpt"]).Trim();
            clsFtp._strLocPath = ((string)rowDiv["tab_pth"]).Trim();
            clsFtp._strDivTip = ((string)rowDiv["tab_tip"]).Trim();

            //ArrayList aryFil = new ArrayList();
            s = ((string)rowDiv["tab_fpt"]).Trim();

            if(((string)rowDiv["tab_tip"]).Trim() == "MIGROSS")
            {
                a = s.Split('|');
                if(a.Length > 1)
                    s = a[0];
            }

            if (!s.Contains('|'))
                s += "|";

            a = s.Split('|');

            if (((string)rowDiv["tab_tip"]).Trim() == "FTPDIV")
            {
                s = ((string)rowDiv["tab_fpt"]).Trim();

                a = s.Split(';');
                if (a.Length > 1)
                    s = a[0];
            }
            
            if(a.Length > 0)
            {
                string sFtpPath = a[0];

                s = a[1];
                a = s.Split(',');

                foreach(string ss in a)
                {
                    clsFtp._strFtpPath = sFtpPath + ss + "/";
                    ArrayList aryFil = clsFtp.FtpRead(ref sMsg, aPre);

                    if (aryFil.Count > 0)
                    {
                        foreach (string sFil in aryFil)
                        {
                            b = true;
                            sMsg += clsFtp.FtpDownLoad(sFil);
                            if (((string)rowDiv["tab_tip"]).Trim() == "FTPDIV")
                                sMsg += clsFtp.FtpRename(sFil, "Old/" + sFil + "_" + DateTime.Now.ToString("yyyyMMddHHmmss"));
                            else if (((string)rowDiv["tab_tip"]).Trim() == "GOTTARDO")
                                sMsg += clsFtp.FtpRename(sFil, "999/" + sFil + "_" + DateTime.Now.ToString("yyyyMMddHHmmss"));
                            else if (((string)rowDiv["tab_tip"]).Trim() == "VEGA")
                                sMsg += clsFtp.FtpDelete(sFil);
                            else
                                sMsg += clsFtp.FtpRename(sFil, "_" + sFil + "_" + DateTime.Now.ToString("yyyyMMddHHmmss"));
                        }
                        if (sMsg == "")
                            sMsg = "Ricevuto " + aryFil.Count.ToString() + " documenti";
                    }
                }
                /**/
                if (((string)rowDiv["tab_tip"]).Trim() == "FTPDIV")
                    _clsQry.ApPhoneDiv2TmpDiv(((string)rowDiv["tab_pth"]).Trim(), aPre);
                else if (((string)rowDiv["tab_tip"]).Trim() == "GOTTARDO")
                    GottardoZip(rowDiv); 
                /**/
            }

            lblDiv.BackColor = bc;
            lblDiv.Text = "";

            return sMsg;
        }

        private void GottardoZip(DataRow rowDiv)
        {
            clsSharpZip cZip = new clsSharpZip();

            string s = "";

            string sPathRead = ((string)rowDiv["tab_pth"]).Trim();
            //string sPathDest = ((string)rowDiv["tab_pth"]).Trim();

            string sVarDest = "";
            string sVarFil = ((string)rowDiv["tab_fil"]).Trim();
            string[] a = sVarFil.Split(',');
            if (a.Length > 1)
            {
                sVarFil = a[0];
                sVarDest = a[1];
            }

            string sFatDest = "";
            string sFatFil = ((string)rowDiv["tab_fat"]).Trim();
            a = sFatFil.Split(',');
            if (a.Length > 1)
            {
                sFatFil = a[0];
                sFatDest = a[1];
            }

            foreach (string sFi in Directory.GetFiles(sPathRead))
            {
                if(Path.GetExtension(sFi).ToUpper() == ".ZIP")
                {
                    string sPathDest = ((string)rowDiv["tab_pth"]).Trim();

                    s = Path.GetFileName(sFi).ToUpper();

                    if (s.Length >= sVarFil.Length && s.Substring(0, sVarFil.Length) == sVarFil.ToUpper() && sVarDest != "")
                        sPathDest = sVarDest;
                    else if (s.Length >= sFatFil.Length && s.Substring(0, sFatFil.Length) == sFatFil.ToUpper() && sFatDest != "")
                        sPathDest = sFatDest;

                    cZip.UnZip(sFi, sPathDest, true);

                    //Spostare il file in old
                    s = Path.GetDirectoryName(sFi) + "\\old\\" + Path.GetFileName(sFi) + "_" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".ZIP";
                    if(File.Exists(s))
                        File.Delete(s);
                    File.Move(sFi, s);

                    Thread.Sleep(2000);

                }
            }
        }

        private void dgv1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.ColumnIndex == 0)
            {
                if (e.RowIndex >= 0)
                {
                    string s = dgv1.Rows[e.RowIndex].Cells[e.ColumnIndex].Value.ToString();

                    if (s == "E")
                        dgv1.Rows[e.RowIndex].Cells[e.ColumnIndex].Value = "P";
                    else if (s == "P")
                        dgv1.Rows[e.RowIndex].Cells[e.ColumnIndex].Value = "C";
                    else if (s == "C")
                        dgv1.Rows[e.RowIndex].Cells[e.ColumnIndex].Value = "E";
                }
            }
        }


        private string CodFromTab(string strTab, string strDes)
        {
            string s = "";
            string sCod = "";

            if (strDes == "MELANZANE VIOLA")
                Console.WriteLine("aaaaaaaaaaa");

            s = "SELECT * FROM " + strTab + " ORDER BY tab_cod DESC";
            DataTable t = _clsFun.FillTabSql(strTab, s, false, _strConSql);

            //if (strDes.Length > 5 && strDes.Substring(0, 5) == "COSTA")
            //    Console.WriteLine("aaaaaaaaaaaa");

            s = _clsFun.FaiLApice(strDes);
            DataView v = new DataView(t, "tab_des='" + s + "'","",DataViewRowState.CurrentRows);

            if(v.Count > 0 )
                sCod = (string)v[0]["tab_cod"];
            else
            {
                if (t.Rows.Count > 0)
                {
                    s = (string)t.Rows[0]["tab_cod"];
                    sCod = (Convert.ToInt16(s) + 1).ToString("000");
                }
                else
                    sCod = "001";

                DataRow x = t.NewRow();
                x["tab_cod"] = sCod;
                x["tab_des"] = strDes;
                x["tab_ann"] = false;

                s = _clsFun.SqlInsertRow(strTab, t, x);

                if (s != "")
                    _clsFun.SqlWrite(s, _strConSql);
            }

            return sCod;
        }

        private void conversioneTabelleToolStripMenuItem_Click(object sender, EventArgs e)
        {
            CurrencyManager cm = dgv1.BindingContext[dgv1.DataSource, dgv1.DataMember] as CurrencyManager;
            if (cm.Position >= 0)
            {
                DataRowView r = cm.List[dgv1.CurrentRow.Index] as DataRowView;
                string sFor = (string)r.Row["dit_for"];

                string p = "TabTabFornitori";
                string s = "SELECT * FROM TabTabFornitori WHERE tab_for='" + sFor + "' ORDER BY tab_tip";
                DataTable t = _clsFun.FillTabSql(p, s, false, _strConSql);
                DataTable tDiv = ((DataView)dgv2.DataSource).ToTable();

                frmGesDivCtrlTab f = new frmGesDivCtrlTab();
                f._strTip = "ALL";
                f._tabTab = t.Copy();
                f._tabDiv = tDiv;
                f._strForCod = sFor;
                f.ShowDialog();
            }
        }

        private void attivaArticoliDelDocumentoToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("Confermi l'attivazione alle casse di tutti gli articoli del documento?", "ATTIVAZIONE ARTICOLI", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                AttArticoli();
        }

        private void AttArticoli()
        {

            string s = "";

            clsVariazioni clsVar = new clsVariazioni();

            DataTable t = ((DataView)dgv2.DataSource).Table;

            progressBar1.Value = 0;
            progressBar1.Maximum = t.Rows.Count;
            progressBar1.Minimum = 0;

            foreach (DataRow y in t.Rows)
            {
                progressBar1.Increment(1);
                Application.DoEvents();

                if ((string)y["div_arf"] == "011594201")
                    Console.WriteLine("xxxx");

                if (((string)y["div_art"]).Trim() != "" && ((string)y["div_sta"]).Trim() != _clsDef.STAATT)
                {
                    s = "UPDATE AnaArticoli SET art_sta='" + _clsDef.STAATT + "' WHERE art_cod='" + y["div_art"] + "'";
                    _clsFun.SqlWrite(s, _strConSql);

                    clsVar.Variazioni((string)y["div_art"], "Forza", "Attivazione da assortimento fornitori", "ALL");

                    y["div_sta"] = _clsDef.STAATT;
                }
                else
                    y["div_sta"] = _clsDef.STAATT;
            }
        }

        private void articoliNuoviToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (dgv2.DataSource != null)
            {
                CurrencyManager cm = dgv1.BindingContext[dgv1.DataSource, dgv1.DataMember] as CurrencyManager;
                if (cm.Position >= 0)
                {
                    DataRowView r = cm.List[dgv1.CurrentRow.Index] as DataRowView;
                    DataRow x = r.Row;

                    string sDes = (string)x["dit_des"];

                    ArticoliNuoviCsv(((DataView)dgv2.DataSource).ToTable(), sDes);
                }
            }
        }

        private void ArticoliNuoviCsv(DataTable tabArt, String strDes)
        {
            DataRow[] j;
            DataRow x;

            DataView v = new DataView(tabArt, "div_art=''", "div_ard", DataViewRowState.CurrentRows);

            DataTable t = v.ToTable();
            DataTable tTmp = new clsGenTabTmp().TabTmpDivArtPrn("TabPrn");

            if(v.Count > 0)
	        {

                DataTable tDie = _dasGen.Tables[TABTMPEAN];
                DataColumn[] keys = new DataColumn[2];
                keys[0] = tDie.Columns["die_arf"];
                keys[1] = tDie.Columns["die_ean"];
                tDie.PrimaryKey = keys;

                foreach(DataRow y in t.Rows)
                {
                    x = tTmp.NewRow();
                    x["tmp_arf"] = y["div_arf"];
                    x["tmp_ard"] = y["div_ard"];
                    x["tmp_cos"] = y["div_cos"];
                    x["tmp_prp"] = y["div_prp"];
                    x["tmp_plu"] = y["div_plu"];
                    x["tmp_ean"] = "";
                    x["tmp_sta"] = "Nuovo";
                    tTmp.Rows.Add(x);

                    j = tDie.Select("die_arf='" + y["div_arf"] + "'");

                    if(j.Length > 0)
                        tTmp.Rows[tTmp.Rows.Count-1]["tmp_ean"] = j[0]["die_ean"];

                    for (int i = 1; i < j.Length; i++)
                    {

                        x = tTmp.NewRow();
                        x["tmp_arf"] = y["div_arf"];
                        //x["tmp_ard"] = y["div_ard"];
                        //x["tmp_cos"] = y["div_cos"];
                        //x["tmp_prp"] = y["div_prp"];
                        //x["tmp_plu"] = y["div_plu"];
                        x["tmp_ean"] = j[i]["die_ean"];
                        //x["tmp_sta"] = "Nuovo";

                        //if (i == 0)
                        //{
                        //    //y["div_ean"] = j[i]["die_ean"];

                        //}
                        //else
                        //if(i > 0)
                        //{


                            //x["div_arf"] = y["div_arf"];
                            //x["div_ard"] = ""; // i.ToString();
                            //x["div_ean"] = j[i]["die_ean"];
                            //x["div_plu"] = y["div_plu"];

                        //x["tmp_arf"] = y["div_arf"];
                        //x["tmp_ard"] = "";
                        //x["tmp_cos"] = y["div_cos"];
                        //x["tmp_prp"] = y["div_prp"];
                        //x["tmp_ean"] = j[i]["die_ean"];
                        //x["tmp_sta"] = "Nuovo";

                        //}

                        tTmp.Rows.Add(x);


                    }
                }

                //foreach(DataRow y in tTmp.Rows)
                //{
                //    t.ImportRow(y);
                //}

                //v = new DataView(t, "div_art=''", "div_arf", DataViewRowState.CurrentRows);

                //t = v.ToTable();

            }

            v = new DataView(tabArt, "div_stf='N'", "div_ard", DataViewRowState.CurrentRows);
            foreach(DataRowView r in v)
            {
                //r["tmp_sta"] = "Annullato";
                //t.ImportRow(r.Row);

                j = tTmp.Select("tmp_arf='" + r["div_arf"] + "'");
                if (j.Length > 0)
                {
                    j[0]["tmp_sta"] = "Annullato";
                }
                else
                {

                    x = tTmp.NewRow();
                    x["tmp_art"] = r["div_art"];
                    x["tmp_arf"] = r["div_arf"];
                    x["tmp_ard"] = r["div_ard"];
                    x["tmp_cos"] = r["div_cos"];
                    x["tmp_prp"] = r["div_prp"];
                    x["tmp_plu"] = r["div_plu"];
                    //x["tmp_ean"] = r["die_ean"];
                    x["tmp_sta"] = "Annullato";
                    tTmp.Rows.Add(x);
                }
            }

            if (t.Rows.Count == 0)
                MessageBox.Show("Non sono presenti nuove/annullate referenze!", "ESTRAZIONE ARTICOLI NUOVI-ANNULLATI", MessageBoxButtons.OK, MessageBoxIcon.Error);
            else
            {

                string sTit = DateTime.Today.ToString("dd/MM/yyyy") + " Articoli nuovi " + strDes;
                string sFil = "ArticoliNuovi";

                string sFld = "";
                sFld += "tmp_art, Articolo, 70, StringLiteral;";
                sFld += "tmp_arf, Articolo fornitore, 70, StringLiteral;";
                sFld += "tmp_ard, Descrizione, 30, StringLiteral;";
                sFld += "tmp_cos, Costo, 90, StringLiteral;";
                sFld += "tmp_prp, P.consigliato, 70, StringLiteral;";
                //sFld += "ean_001, EAN, 70, StringLiteral;";
                //sFld += "ean_002, EAN, 70, StringLiteral;";
                //sFld += "ean_003, EAN, 70, StringLiteral;";
                //sFld += "ean_004, EAN, 70, StringLiteral;";
                //sFld += "ean_005, EAN, 70, StringLiteral;";
                sFld += "tmp_plu, PLU, 70, StringLiteral;";
                sFld += "tmp_ean, EAN, 70, StringLiteral;";
                sFld += "tmp_sta, Stato, 70, StringLiteral;";

                (new clsExcel()).exportToCsv1(sFld, tTmp, sFil, sTit, "");
            }

        }

        private string DitNewNum()
        {
            string s = "001";

            if (dgv1.DataSource != null)
            {
                DataView v = new DataView((DataTable)dgv1.DataSource, "", "dit_num DESC", DataViewRowState.CurrentRows);
                if (v.Count > 0)
                    s = (Convert.ToInt16(v[0]["dit_num"]) + 1).ToString("000");
            }

            return s;
        }

        private void differenzeToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (dgv2.DataSource != null)
            {
                CurrencyManager cm = dgv1.BindingContext[dgv1.DataSource, dgv1.DataMember] as CurrencyManager;
                if (cm.Position >= 0)
                {
                    DataRowView r = cm.List[dgv1.CurrentRow.Index] as DataRowView;
                    DataRow x = r.Row;

                    string sFor = (string)x["dit_for"];
                    string sDes = (string)x["dit_des"];

                    ArticoliDiffAnagrafiche(((DataView)dgv2.DataSource).ToTable(), sFor, sDes);
                }
            }
        }

        private void stampaDocumentoToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (dgv2.DataSource != null)
            {
                CurrencyManager cm = dgv1.BindingContext[dgv1.DataSource, dgv1.DataMember] as CurrencyManager;
                if (cm.Position >= 0)
                {
                    DataRowView r = cm.List[dgv1.CurrentRow.Index] as DataRowView;
                    DataRow x = r.Row;

                    string sFor = (string)x["dit_for"];

                    if (!DBNull.Value.Equals(x["dit_tfo"]))
                        sFor += " " + (string)x["dit_tfo"];

                    string sDes = (string)x["dit_des"] + " del " + ((DateTime)x["dit_ddo"]).ToString("dd/MM/yyyy");

                    new clsStampe().PrintDivFornitore(((DataView)dgv2.DataSource).ToTable(), sFor, sDes);
                }
            }
        }

        private void ArticoliDiffAnagrafiche(DataTable tabArt, String strFor, String strDes)
        {
            DataRow[] j;
            DataRow[] j2;
            DataRow x;
            DataRow x2;
            Boolean b = false;

             //_strForRif == strFor

            Color bc = lblDiv.BackColor;
            lblDiv.BackColor = Color.LightCoral;
            lblDiv.Text = "Stampa in corso";
            System.Windows.Forms.Application.DoEvents();

            string s = "SELECT * FROM AnaArticoli";
            DataTable tArt = _clsFun.FillTabSql(TABANAART, s, false, _strConSql);
            DataColumn[] keys = new DataColumn[1];
            keys[0] = tArt.Columns["art_cod"];
            tArt.PrimaryKey = keys;

            DataTable tTmp = new clsGenTabTmp().TabTmpDivArtDif("ArtTmp");
            keys = new DataColumn[2];
            keys[0] = tTmp.Columns["tmp_art"];
            keys[1] = tTmp.Columns["tmp_tip"];
            tTmp.PrimaryKey = keys;

            DataView v = new DataView(tabArt, "", "div_art", DataViewRowState.CurrentRows);

            //DataTable tDie = _dasGen.Tables[TABTMPEAN];
            //DataColumn[] keys = new DataColumn[2];
            //keys[0] = tDie.Columns["die_arf"];
            //keys[1] = tDie.Columns["die_ean"];
            //tDie.PrimaryKey = keys;

            DataTable t = v.ToTable();

            progressBar1.Value = 0;
            progressBar1.Maximum = t.Rows.Count;
            progressBar1.Minimum = 0;

            foreach(DataRow y in t.Rows)
            {
                progressBar1.Increment(1);
                System.Windows.Forms.Application.DoEvents();

                if ((string)y["div_art"] == "0007888")
                    Console.WriteLine("aaaa");

                if ((string)y["div_art"] != "")
                {
                    j = tArt.Select("art_cod='" + y["div_art"] + "'");
                    if (j.Length > 0)
                    {
                        j2 = tTmp.Select("tmp_tip='VAR' AND tmp_art='" + y["div_art"] + "'");
                        if (j2.Length == 0)
                        {
                            x = tTmp.NewRow();
                            x["tmp_art"] = y["div_art"];
                            x["tmp_tip"] = "VAR";
                            x["tmp_div"] = "DIV";

                            b = false;

                            x["tmp_ard"] = "";
                            if ((string)y["div_ard"] != (string)j[0]["art_des"])
                            {
                                x["tmp_ard"] = y["div_ard"];
                                if (_strForRif == strFor)
                                    b = true;
                            }
                            x["tmp_umi"] = "";
                            if ((string)y["div_umi"] != (string)j[0]["art_umi"])
                            {
                                x["tmp_umi"] = y["div_umi"];
                                b = true;
                            }
                            x["tmp_pne"] = 0;
                            if (Convert.ToDecimal(y["div_pne"]) != Convert.ToDecimal(j[0]["art_pne"]))
                            {
                                x["tmp_pne"] = y["div_pne"];
                                b = true;
                            }
                            x["tmp_pxc"] = 0;
                            if (Convert.ToDecimal(y["div_pxc"]) != Convert.ToDecimal(j[0]["art_pxc"]))
                            {
                                x["tmp_pxc"] = y["div_pxc"];
                                b = true;
                            }
                            x["tmp_plu"] = "";

                            if (_clsFun.Numerico((string)j[0]["art_plu"]))
                                j[0]["art_plu"] = Convert.ToInt16(j[0]["art_plu"]).ToString();

                            if ((string)y["div_plu"] != (string)j[0]["art_plu"])
                            {
                                x["tmp_plu"] = y["div_plu"];
                                b = true;
                            }
                            x["tmp_rep"] = "";
                            if ((string)y["div_rep"] != (string)j[0]["art_rep"])
                            {
                                x["tmp_rep"] = y["div_rep"];
                                b = true;
                            }
                            x["tmp_ori"] = "";
                            if ((string)y["div_ori"] != (string)j[0]["art_ori"])
                            {
                                x["tmp_ori"] = y["div_ori"];
                                b = true;
                            }
                            x["tmp_cat"] = "";
                            if ((string)y["div_cat"] != (string)j[0]["art_cat"])
                            {
                                x["tmp_cat"] = y["div_cat"];
                                b = true;
                            }
                            if (b)
                            {
                                tTmp.Rows.Add(x);
                                x2 = tTmp.NewRow();
                                x2["tmp_art"] = (string)x["tmp_art"];
                                x2["tmp_tip"] = "ART";

                                //if ((string)y["div_ard"] != (string)j[0]["art_des"])
                                x2["tmp_ard"] = (string)j[0]["art_des"];
                                if ((string)y["div_umi"] != (string)j[0]["art_umi"])
                                    x2["tmp_umi"] = (string)j[0]["art_umi"];
                                if ((decimal)y["div_pne"] != (decimal)j[0]["art_pne"])
                                    x2["tmp_pne"] = (decimal)j[0]["art_pne"];
                                if (Convert.ToDecimal(y["div_pxc"]) != Convert.ToDecimal(j[0]["art_pxc"]))
                                    x2["tmp_pxc"] = Convert.ToDecimal(j[0]["art_pxc"]);
                                if ((string)y["div_plu"] != (string)j[0]["art_plu"])
                                    x2["tmp_plu"] = (string)j[0]["art_plu"];
                                if ((string)y["div_rep"] != (string)j[0]["art_rep"])
                                    x2["tmp_rep"] = (string)j[0]["art_rep"];
                                if ((string)y["div_reb"] != (string)j[0]["art_reb"])
                                    x2["tmp_reb"] = (string)j[0]["art_reb"];
                                if ((string)y["div_ori"] != (string)j[0]["art_ori"])
                                    x2["tmp_ori"] = (string)j[0]["art_ori"];
                                if ((string)y["div_cat"] != (string)j[0]["art_cat"])
                                    x2["tmp_cat"] = (string)j[0]["art_cat"];

                                tTmp.Rows.Add(x2);
                            }
                        }
                    }

                }

            }

            //foreach(DataRow y in tTmp.Rows)
            //{
            //    t.ImportRow(y);
            //}

            v = new DataView(tTmp, "", "tmp_art, tmp_tip", DataViewRowState.CurrentRows);
            t = v.ToTable();

            string sTit = DateTime.Today.ToString("dd/MM/yyyy") + " Deffeneze anagrafiche " + strDes;
            string sFil = "DiffAnagrafiche";

            string sFld = "";
            sFld += "tmp_div, T., 70, StringLiteral;";
            sFld += "tmp_art, Articolo, 70, StringLiteral;";
            sFld += "tmp_tip, Tipo, 70, StringLiteral;";
            sFld += "tmp_ard, Descrizione, 30, StringLiteral;";
            sFld += "tmp_rep, Reparto, 30, StringLiteral;";
            sFld += "tmp_umi, UMI, 30, StringLiteral;";
            sFld += "tmp_pne, P.netto, 30, StringLiteral;";
            sFld += "tmp_pxc, PxC, 30, StringLiteral;";
            sFld += "tmp_reb, R.Bil., 30, StringLiteral;";
            sFld += "tmp_plu, PLU, 30, StringLiteral;";
            sFld += "tmp_ori, Origine, 30, StringLiteral;";
            sFld += "tmp_cat, Categoria, 30, StringLiteral;";

            (new clsExcel()).exportToCsv1(sFld, t, sFil, sTit, "");

            lblDiv.BackColor = bc;
            lblDiv.Text = "";

	    }

        private void DivNewFornitori(string strTra, string strFor, string strFod)
        {
            DataTable tTes = (DataTable)dgv1.DataSource;
            DataTable tRig = _dasGen.Tables[TABTMPFDR];

            //DataTable tEan = _dasGen.Tables[TABTMPEAN];
            //DataColumn[] keys = new DataColumn[2];
            //keys[0] = tEan.Columns["die_arf"];
            //keys[1] = tEan.Columns["die_ean"];
            //tEan.PrimaryKey = keys;


            DataTable tDie = _dasGen.Tables[TABTMPEAN];
            DataColumn[] keys = new DataColumn[2];
            keys[0] = tDie.Columns["die_arf"];
            keys[1] = tDie.Columns["die_ean"];
            tDie.PrimaryKey = keys;

            string s = "SELECT * FROM AnaForDivTestate WHERE trt_for='" + strTra + "'";
            //DataTable tTrt = _clsFun.FillTabSql("AnaForDivTestate", s, false, _strConSql);
            DataTable tTrt = _clsFun.FillTabMdb("AnaForDivTestate", s, false, _strConSqlMdb);

            DataRow[] j;
            DataRow x;
            string sNum = "";
            string sNums = "";
            string sNdo = "";
            DateTime dDdo = DateTime.Today;
            string sTva = "";

            foreach (DataRow y in tTrt.Rows)
            {
                if ((string)y["trt_tip"] == "ART" || (string)y["trt_tip"] == "FAT" || (string)y["trt_tip"] == "OFF")
                {
                    string sPth = ((string)y["trt_pth"]).Trim();

                    //string sPathFor = ((string)y["tab_pth"]).Trim();

                    s = ((string)y["trt_fil"]).Trim();
                    string[] aFil = s.Split(',');

                    string[] sFils = Directory.GetFiles(sPth);
                    foreach (string sFi in sFils)
                    {

                        string sFil = sFi;

                        Boolean b = false;

                        foreach (string ss in aFil)
                        {
                            s = Path.GetFileName(sFil);

                            if (ss != "" && s.Length >= ss.Length && s.Substring(0, ss.Length).ToLower() == ss.ToLower())
                            {
                                b = true;
                                break;
                            }
                        }

                        if (b)
                        {
                            string sLeg = ((string)y["trt_leg"]).Trim();             //Archivio legato: ART + EAN

                            sTva = ((string)y["trt_tip"]).Trim();
                            int iRig = 0;

                            FileInfo fInfo = new FileInfo(sFil);

                            progressBar1.Value = 0;
                            progressBar1.Maximum = (int)fInfo.Length;
                            progressBar1.Minimum = 0;

                            using (StreamReader sr = new StreamReader(sFil, System.Text.Encoding.Default))
                            {
                                string sRig = "";

                                s = "SELECT * FROM AnaForDivDettaglio WHERE tra_for='" + strTra + "' AND tra_cod='" + (string)y["trt_cod"] + "' AND tra_ann=0 ORDER BY tra_pos";
                                DataTable tTra = _clsFun.FillTabMdb("AnaForDivDettaglio", s, false, _strConSqlMdb);

                                while ((sRig = sr.ReadLine()) != null)
                                {
                                    progressBar1.Increment(sRig.Length);
                                    System.Windows.Forms.Application.DoEvents();

                                    if (sRig.Length > 20)
                                    {
                                        iRig++;

                                        x = tRig.NewRow();
                                        x["div_dva"] = DateTime.Today;

                                        //if (iRig == 59)
                                        //    Console.WriteLine("xxxxx");

                                        foreach (DataRow xx in tTra.Rows)
                                        {
                                            s = (string)xx["tra_fld"];
                                            if (s == "art_prp")
                                                Console.WriteLine("zzzzzzzzz");

                                            DivNewRow(sRig, xx, x);
                                        }

                                        if (!DBNull.Value.Equals(x["div_ean"]) && ((string)x["div_ean"]).Trim() != "")
                                        {
                                            string sBar = ((string)x["div_ean"]).Trim();

                                            for (int ii = 0; ii < sBar.Length; ii += 13)
                                            {
                                                //s = sBar.Substring(ii, 13);

                                                if (ii + 13 > sBar.Length)
                                                    s = sBar.Substring(ii);
                                                else
                                                    s = sBar.Substring(ii, 13);

                                                s = s.Trim();

                                                if (s.Trim() != "" && s.Length > 7 && _clsFun.Numerico(s))
                                                {
                                                    s = s.Substring(0, s.Length - 1) + new clsCtrlCodici().FindMod10Digit(s.Substring(0, s.Length - 1));
                                                    j = tDie.Select("die_arf='" + x["div_arf"] + "' AND die_ean='" + s + "'");
                                                    if (j.Length == 0)
                                                    {
                                                        if (s == "0178")
                                                            Console.WriteLine("aaaaaaaaaa");

                                                        DataRow yy = tDie.NewRow();
                                                        yy["die_num"] = sNum;
                                                        yy["die_for"] = strFor;
                                                        yy["die_arf"] = x["div_arf"];
                                                        yy["die_ean"] = s;
                                                        tDie.Rows.Add(yy);

                                                        if (s != ((string)x["div_ean"]).Trim())
                                                            _clsFun.ErrorLog("Divulgazione DADO", "Corretto check digit " + ((string)x["div_ean"]).Trim() + " -> " + s);

                                                    }
                                                }
                                            }

                                            x["div_ean"] = "";

                                        }

                                        if (iRig == 1)
                                        {
                                            j = tTes.Select("dit_tva='" + (string)y["trt_tip"] + "' AND dit_ndo='" + sNdo + "' AND dit_ddo=#" + dDdo.ToString("MM/dd/yyyy") + "#");
                                            if (j.Length == 0)
                                            {
                                                //sNum = (tTes.Rows.Count + 1).ToString("000");
                                                sNum = DitNewNum();

                                                sNums += sNum + ";";
                                                DataRow xx = tTes.NewRow();
                                                xx["dit_num"] = sNum;
                                                xx["dit_for"] = strFor;
                                                xx["dit_des"] = strFod;
                                                xx["dit_fil"] = Path.GetFileName(sFil);
                                                xx["dit_tva"] = sTva;

                                                xx["dit_ndo"] = "";
                                                if ((string)y["trt_tip"] == "FAT")
                                                {
                                                    xx["dit_ndo"] = x["div_num"];
                                                    xx["dit_ddo"] = x["div_dva"];
                                                }

                                                xx["dit_ddo"] = dDdo;
                                                xx["dit_pth"] = Path.GetDirectoryName(sFil);
                                                xx["dit_nri"] = 1;
                                                xx["dit_cho"] = "E";

                                                if (sTva == "OFF")
                                                {
                                                    xx["dit_soi"] = x["off_soi"];
                                                    xx["dit_sof"] = x["off_soo"];
                                                    xx["dit_ndo"] = x["div_off"];

                                                }

                                                tTes.Rows.Add(xx);
                                                iRig = 1;
                                                dDdo = (DateTime)xx["dit_ddo"];
                                            }
                                        }

                                        x["div_tva"] = sTva;
                                        x["div_num"] = sNum;
                                        x["div_nri"] = iRig.ToString("00000");


                                        tRig.Rows.Add(x);
                                    }
                                }
                            }

                            if (sLeg != "")
                            {
                                string[] a = sLeg.Split('-');
                                if (a.Length > 1)
                                {
                                    string sTip = a[0];
                                    sFil = a[1];

                                    j = tTrt.Select("trt_tip='" + sTip + "' AND trt_fil='" + sFil + "'");
                                    if (j.Length > 0 && sTip == "EAN")
                                    {
                                        sPth = ((string)j[0]["trt_pth"]).Trim();
                                        sFil = sPth + sFil;
                                        string sCod = ((string)j[0]["trt_cod"]).Trim();

                                        if (File.Exists(sFil))
                                        {

                                            FileInfo fInfoLeg = new FileInfo(sFil);

                                            progressBar1.Value = 0;
                                            progressBar1.Maximum = (int)fInfoLeg.Length;
                                            progressBar1.Minimum = 0;

                                            using (StreamReader sr = new StreamReader(sFil, System.Text.Encoding.Default))
                                            {
                                                string sRig = "";

                                                s = "SELECT * FROM AnaForDivDettaglio WHERE tra_for='" + strTra + "' AND tra_cod='" + sCod + "' AND tra_ann=0 ORDER BY tra_pos";
                                                //DataTable tTra = _clsFun.FillTabSql("AnaForDivDettaglio", s, false, _strConSql);
                                                DataTable tTra = _clsFun.FillTabMdb("AnaForDivDettaglio", s, false, _strConSqlMdb);

                                                while ((sRig = sr.ReadLine()) != null)
                                                {
                                                    progressBar1.Increment(sRig.Length);
                                                    System.Windows.Forms.Application.DoEvents();

                                                    x = tRig.NewRow();

                                                    foreach (DataRow xx in tTra.Rows)
                                                    {
                                                        DivNewRow(sRig, xx, x);
                                                    }

                                                    string sBar = "";
                                                    if (!DBNull.Value.Equals(x["div_ean"]) && ((string)x["div_ean"]).Trim() != "")
                                                    {
                                                        sBar = ((string)x["div_ean"]).Trim();

                                                        for (int ii = 0; ii < sBar.Length; ii += 13)
                                                        {
                                                            //s = sBar.Substring(ii, 13);

                                                            if (ii + 13 > sBar.Length)
                                                                s = sBar.Substring(ii);
                                                            else
                                                                s = sBar.Substring(ii, 13);

                                                            s = s.Trim();

                                                            if (s != "" && s.Length > 7 && _clsFun.Numerico(s))
                                                            {
                                                                s = s.Substring(0, s.Length - 1) + new clsCtrlCodici().FindMod10Digit(s.Substring(0, s.Length - 1));
                                                                j = tDie.Select("die_arf='" + x["div_arf"] + "' AND die_ean='" + s + "'");
                                                                if (j.Length == 0)
                                                                {
                                                                    if (s == "0178")
                                                                        Console.WriteLine("aaaaaaaaaa");

                                                                    DataRow yy = tDie.NewRow();
                                                                    yy["die_num"] = sNum;
                                                                    yy["die_for"] = strFor;
                                                                    yy["die_arf"] = x["div_arf"];
                                                                    yy["die_ean"] = s;
                                                                    tDie.Rows.Add(yy);

                                                                    if (s != ((string)x["div_ean"]).Trim())
                                                                        _clsFun.ErrorLog("Divulgazione DADO", "Corretto check digit " + ((string)x["div_ean"]).Trim() + " -> " + s);

                                                                }
                                                            }
                                                        }
                                                    }

                                                    x["div_ean"] = "";

                                                }
                                            }
                                        }
                                    }
                                }
                            }

                            string[] aa = sNums.Split(';');
                            for (int i = 0; i < aa.Length; i++)
                            {
                                DataTable tTab = CtrlDivArt(sFil, tRig, strFor, aa[i]);
                                DataView v = new DataView(tTab, "tab_cod=''", "", DataViewRowState.CurrentRows);
                                if (v.Count > 0)
                                {
                                    frmGesDivCtrlTab f = new frmGesDivCtrlTab();
                                    f._tabTab = tTab.Copy();
                                    f._strForCod = strFor;
                                    f.ShowDialog();
                                    if (f._bolMdy)
                                    {
                                        CtrlDivArt(sFil, tRig, strFor, aa[i]);
                                    }
                                }
                            }
                        }
                    }
                }
            }
        }

        private void DivNewRow(string strRig, DataRow rowRig, DataRow rowRow)
        {
            string s = "";
            string sFld = (string)rowRig["tra_fld"];
            int iPos = Convert.ToInt16(rowRig["tra_pos"]);
            int iLen = Convert.ToInt16(rowRig["tra_len"]);
            string sReg = ((string)rowRig["tra_reg"]).Trim();

            if (strRig.Length >= iPos + iLen)
            {
                string sVal = strRig.Substring(iPos, iLen).Trim();
                if (sVal != "")
                {
                    if (sReg.Length > 2 && sReg.Substring(0, 2) == "IF")
                    {
                        s = sReg.Substring(3);
                        string[] a = s.Split('=');
                        if (a.Length > 2)
                        {
                            if (sVal == a[0])
                                sVal = a[1];
                            else
                                sVal = a[2];
                            rowRow[sFld] = sVal;
                        }
                    }
                    else if (sReg == "AAAAMMDD")
                        rowRow[sFld] = _clsFun.Str2Day(sVal);
                    else if (sReg == "INT")
                        rowRow[sFld] = Convert.ToInt16(sVal);
                    else if (sReg.Length > 3 && sReg.Substring(0, 3) == "DEC")
                    {
                        if (_clsFun.Numerico(sVal))
                        {
                            int dDec = Convert.ToInt16(sReg.Substring(4));
                            decimal dVal = Convert.ToDecimal(sVal);
                            if (dDec > 0 && dVal > 0)
                                rowRow[sFld] = dVal / dDec;
                        }
                    }
                    else if (sReg.Length > 5 && sReg.Substring(0, 4) == "PADL")
                    {
                        if (_clsFun.Numerico(sVal))
                        {
                            string sFill = sReg.Substring(5);

                            s = sFill + sVal.Trim();

                            sVal = s.Substring(s.Length - sFill.Length, sFill.Length);
                            rowRow[sFld] = sVal;
                        }
                    }
                    else
                        rowRow[sFld] = sVal;
                }
            }
        }

        private void DivNewFornitori2(string strTra, string strFor, string strFod)
        {
            DataTable tTes = (DataTable)dgv1.DataSource;
            DataTable t = _dasGen.Tables[TABTMPFDR];

            DataTable tDie = _dasGen.Tables[TABTMPEAN];
            DataColumn[] keys = new DataColumn[2];
            keys[0] = tDie.Columns["die_arf"];
            keys[1] = tDie.Columns["die_ean"];
            tDie.PrimaryKey = keys;

            string s = "SELECT * FROM AnaForDivTestate WHERE trt_for='" + strTra + "'";
            //DataTable tTrt = _clsFun.FillTabSql("AnaForDivTestate", s, false, _strConSql);
            DataTable tTrt = _clsFun.FillTabMdb("AnaForDivTestate", s, false, _strConSqlMdb);

            DataRow[] j;
            DataRow x;
            string sNum = "";
            string sNums = "";
            DateTime dDdo = DateTime.Today;

            if (strFod == "PENGO")
                Console.WriteLine("aaaaaaaa");

            //Ricerco il File/Tracciato
            s = DivTraSeek(strTra);

            if (s != "")
            {
                string[] aa = s.Split('-');

                foreach (string ss in aa)
                {
                    string[] a = ss.Split('+');

                    if (a.Length > 1)
                    {
                        string sTrtTip = a[0];
                        string sTrtFil = a[1];
                        string sTrtLeg = a[2].Trim();

                        s = "SELECT ";
                        s += "AnaForDivTestate.trt_for, ";
                        s += "AnaForDivTestate.trt_cod, ";
                        s += "AnaForDivTestate.trt_fil, ";
                        s += "AnaForDivTestate.trt_pth, ";
                        s += "AnaForDivTestate.trt_ann, ";
                        s += "AnaForDivTestate.trt_tip, ";
                        s += "AnaForDivTestate.trt_leg, ";
                        s += "AnaForDivTestate.trt_reg, ";
                        s += "AnaForDivDettaglio.tra_fld, ";
                        s += "AnaForDivDettaglio.tra_des, ";
                        s += "AnaForDivDettaglio.tra_pos, ";
                        s += "AnaForDivDettaglio.tra_len, ";
                        s += "AnaForDivDettaglio.tra_reg, ";
                        s += "AnaForDivDettaglio.tra_ann ";
                        s += "FROM AnaForDivTestate ";
                        s += "INNER JOIN AnaForDivDettaglio ON AnaForDivTestate.trt_for = AnaForDivDettaglio.tra_for AND AnaForDivTestate.trt_cod = AnaForDivDettaglio.tra_cod ";
                        s += "WHERE (";
                        s += "AnaForDivTestate.trt_ann = 0 AND ";
                        s += "AnaForDivDettaglio.tra_ann = 0 AND ";
                        s += "AnaForDivTestate.trt_for = '" + strTra + "' AND ";
                        s += "(AnaForDivTestate.trt_tip = '" + sTrtTip + "' ";
                        if (sTrtLeg != "")
                            s += " OR AnaForDivTestate.trt_tip = 'EAN' ";
                        s += ")) ";
                        s += "ORDER BY AnaForDivTestate.trt_cod, AnaForDivDettaglio.tra_fld";
                        //DataTable tTra = _clsFun.FillTabSql("AnaForDivDettaglio", s, false, _strConSql);
                        DataTable tTra = _clsFun.FillTabMdb("AnaForDivDettaglio", s, false, _strConSqlMdb);

                        for (int iGiro = 1; iGiro <= 2; iGiro++)
                        {
                            if (iGiro == 2)
                            {
                                if (sTrtLeg.Trim() != "")
                                {
                                    a = sTrtLeg.Split('-');
                                    sTrtTip = a[0];
                                    sTrtFil = Path.GetDirectoryName(sTrtFil) + "\\" + a[1];
                                }
                                else
                                    break;
                            }

                            if (Path.GetExtension(sTrtFil).ToLower() == ".xlsx")
                            {
                                s = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + sTrtFil + ";Extended Properties=\"Excel 12.0;HDR={1};IMEX=1\"";

                                OleDbConnection cn = new System.Data.OleDb.OleDbConnection(s);
                                cn.Open();

                                DataTable sch = cn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, new object[] { null, null, null, "TABLE" });
                                string sTab = (string)sch.Rows[0][2];

                                OleDbCommand cm = new OleDbCommand();
                                cm.Connection = cn;

                                s = "SELECT * FROM [" + sTab + "]";
                                cm.CommandText = s;
                                OleDbDataReader dr = cm.ExecuteReader();
                                DataTable tXls = new DataTable();
                                tXls.TableName = "TabXls";
                                tXls.Load(dr);

                                progressBar1.Value = 0;
                                progressBar1.Maximum = tXls.Rows.Count;
                                progressBar1.Minimum = 0;

                                string sRes = "";
                                string sRig = "";
                                int iRig = 0;

                                int iRowIni = 0;
                                s = ForDivRegola("ROWINI", (string)tTrt.Rows[0]["trt_reg"]);
                                if (_clsFun.Numerico(s, "0123456789"))
                                    iRowIni = Convert.ToInt16(s);

                                foreach (DataRow yy in tXls.Rows)
                                {
                                    sRig = "";
                                    iRig++;

                                    foreach (DataColumn c in tXls.Columns)
                                    {
                                        sRig += yy[c.ColumnName] + ";";
                                    }

                                    progressBar1.Increment(1);
                                    System.Windows.Forms.Application.DoEvents();

                                    if (sRig.Length > 20  && iRig >= iRowIni && sRig.Substring(0,10) != ";;;;;;;;;;")
                                    {
                                        s = sTrtTip + "|" + sTrtFil + "|" + strFor + "|" + iRig.ToString("00000") + "|" + strFod;

                                        string sEcr = "";

                                        if (strFod == "VENFRI" && sTrtTip == "ART")
                                            sEcr = TraForEcr(sRig, strFor);

                                        sRes = DivTraRow(sRig, tTra, tTes, t, tDie, s, sRes, sEcr);

                                        a = sRes.Split('-');
                                        if (a.Length > 0)
                                        {
                                            s = a[0];

                                            if (s.Length > 0)
                                            {
                                                if (!sNums.Contains(s))
                                                    sNums += s + ";";
                                            }
                                        }
                                    }
                                }

                                cm.Dispose();
                                cn.Close();
                                cn.Dispose();
                            }
                            else if (Path.GetExtension(sTrtFil).ToLower() == ".csv")
                            {
                                using (StreamReader sr = new StreamReader(sTrtFil, System.Text.Encoding.Default))
                                {
                                    string sRes = "";
                                    string sRig = "";
                                    int iRig = 0;

                                    int iRowIni = 0;
                                    s = ForDivRegola("ROWINI", (string)tTrt.Rows[0]["trt_reg"]);
                                    if (_clsFun.Numerico(s, "0123456789"))
                                        iRowIni = Convert.ToInt16(s);

                                    FileInfo fInfo = new FileInfo(sTrtFil);

                                    progressBar1.Value = 0;
                                    progressBar1.Maximum = (int)fInfo.Length;
                                    progressBar1.Minimum = 0;

                                    while ((sRig = sr.ReadLine()) != null)
                                    {
                                        iRig++;
                                        progressBar1.Increment(sRig.Length);
                                        System.Windows.Forms.Application.DoEvents();

                                        if (sRig.Length > 20 && iRig >= iRowIni)
                                        {
                                            s = sTrtTip + "|" + sTrtFil + "|" + strFor + "|" + iRig.ToString("00000") + "|" + strFod;

                                            sRes = DivTraRow(sRig, tTra, tTes, t, tDie, s, sRes, "");

                                            a = sRes.Split('-');
                                            if (a.Length > 0)
                                            {
                                                s = a[0];

                                                if (s.Length > 0)
                                                {
                                                    if (!sNums.Contains(s))
                                                        sNums += s + ";";
                                                }
                                            }
                                        }
                                    }

                                    sr.Close();
                                    sr.Dispose();
                                }

                            }
                            else
                            {
                                using (StreamReader sr = new StreamReader(sTrtFil, System.Text.Encoding.Default))
                                {
                                    string sRes = "";
                                    string sRig = "";
                                    int iRig = 0;


                                    FileInfo fInfo = new FileInfo(sTrtFil);

                                    progressBar1.Value = 0;
                                    progressBar1.Maximum = (int)fInfo.Length;
                                    progressBar1.Minimum = 0;

                                    while ((sRig = sr.ReadLine()) != null)
                                    {
                                        progressBar1.Increment(sRig.Length);
                                        System.Windows.Forms.Application.DoEvents();

                                        if (sRig.Length > 20)
                                        {
                                            s = sTrtTip + "|" + sTrtFil + "|" + strFor + "|" + iRig.ToString("00000") + "|" + strFod;

                                            sRes = DivTraRow(sRig, tTra, tTes, t, tDie, s, sRes, "");

                                            a = sRes.Split('-');
                                            if (a.Length > 0)
                                            {
                                                s = a[0];

                                                if (s.Length > 0)
                                                {
                                                    if (!sNums.Contains(s))
                                                        sNums += s + ";";
                                                }
                                            }
                                        }
                                    }

                                    sr.Close();
                                    sr.Dispose();
                                }
                            }
                        }

                        a = sNums.Split(';');

                        for (int i = 0; i < a.Length; i++)
                        {
                            if (a[i] != "")
                            {
                                DataTable tTab = CtrlDivArt(sTrtFil, t, strFor, a[i]);
                                DataView v = new DataView(tTab, "tab_cod=''", "", DataViewRowState.CurrentRows);
                                if (v.Count > 0)
                                {
                                    frmGesDivCtrlTab f = new frmGesDivCtrlTab();
                                    f._tabTab = tTab.Copy();
                                    f._tabDiv = t.Copy();
                                    f._strForCod = strFor;
                                    f.ShowDialog();
                                    if (f._bolMdy)
                                    {
                                        CtrlDivArt(sTrtFil, t, strFor, a[i]);
                                    }
                                }
                            }
                        }
                    }
                }
            }
        }

        private string TraForEcr(string strRig, string strFor)
        {
            DataRow x;
            string s = "";
            string sEcr = "";

            string sEc1 = "";
            string sEc2 = "";
            string sEc3 = "";

            string[] a = strRig.Split(';');

            if (a.Length > 10)
            {
                string sLc1 = a[2];
                string sLd1 = a[3];
                string sLc2 = a[4];
                string sLd2 = a[5];
                string sLc3 = a[6];
                string sLd3 = a[7];

                if (sLc2 == "15020")
                    Console.WriteLine("xxx");

                sEcr = sLc1.Trim() + sLc2.Trim() + sLc3.Trim();

                if (sEcr.Length == 15)
                {
                    //Ecr livello 1

                    s = "SELECT * FROM TabTabFornitori WHERE ";
                    s += "tab_tip='EC1' AND ";
                    s += "tab_for='" + strFor + "' AND ";
                    s += "tab_cof='" + sLc1 + "'";
                    DataTable t = _clsFun.FillTabSql("TabTabFornitori", s, true, _strConSql);

                    if (t.Rows.Count == 0)
                    {
                        sEc1 = "001";
                        s = "SELECT * FROM TabEcrLv1 ORDER BY tab_cod DESC";
                        DataTable t2 = _clsFun.FillTabSql("TabEcrLv1", s, true, _strConSql);
                        if (t2.Rows.Count > 0)
                            sEc1 = (Convert.ToInt16(t2.Rows[0]["tab_cod"]) + 1).ToString("000");

                        x = t2.NewRow();
                        x["tab_cod"] = sEc1;
                        x["tab_des"] = sLd1;
                        x["tab_ann"] = false;
                        s = _clsFun.SqlInsertRow("TabEcrLv1", t2, x);
                        _clsFun.SqlWrite(s, _strConSql);

                        x = t.NewRow();
                        x["tab_for"] = strFor;
                        x["tab_tip"] = "EC1";
                        x["tab_cof"] = sLc1;
                        x["tab_cod"] = sEc1;

                        s = _clsFun.SqlInsertRow("TabTabFornitori", t, x);
                        _clsFun.SqlWrite(s, _strConSql);

                    }
                    else
                        sEc1 = (string)t.Rows[0]["tab_cod"];

                    //Ecr livello 2

                    s = "SELECT * FROM TabTabFornitori WHERE ";
                    s += "tab_tip='EC2' AND ";
                    s += "tab_for='" + strFor + "' AND ";
                    s += "tab_cof='" + sLc1 + sLc2 + "'";
                    t = _clsFun.FillTabSql("TabTabFornitori", s, true, _strConSql);

                    if (t.Rows.Count == 0)
                    {
                        sEc2 = "001";
                        s = "SELECT * FROM TabEcrLv2 WHERE tab_lv1='" + sEc1 + "' ORDER BY tab_cod DESC";
                        DataTable t2 = _clsFun.FillTabSql("TabEcrLv2", s, true, _strConSql);
                        if (t2.Rows.Count > 0)
                            sEc2 = (Convert.ToInt16(t2.Rows[0]["tab_cod"]) + 1).ToString("000");

                        x = t2.NewRow();
                        x["tab_cod"] = sEc2;
                        x["tab_lv1"] = sEc1;
                        x["tab_des"] = sLd2;
                        x["tab_ann"] = false;
                        s = _clsFun.SqlInsertRow("TabEcrLv2", t2, x);
                        _clsFun.SqlWrite(s, _strConSql);

                        x = t.NewRow();
                        x["tab_for"] = strFor;
                        x["tab_tip"] = "EC2";
                        x["tab_cof"] = sLc1 + sLc2;
                        x["tab_cod"] = sEc2;

                        s = _clsFun.SqlInsertRow("TabTabFornitori", t, x);
                        _clsFun.SqlWrite(s, _strConSql);
                    }
                    else
                        sEc2 = (string)t.Rows[0]["tab_cod"];


                    //Ecr livello 3

                    s = "SELECT * FROM TabTabFornitori WHERE ";
                    s += "tab_tip='EC3' AND ";
                    s += "tab_for='" + strFor + "' AND ";
                    s += "tab_cof='" + sLc1 + sLc2 + sLc3 + "'";
                    t = _clsFun.FillTabSql("TabTabFornitori", s, true, _strConSql);

                    if (t.Rows.Count == 0)
                    {
                        sEc3 = "001";
                        s = "SELECT * FROM TabEcrLv3 WHERE tab_lv1='" + sEc1 + "' AND tab_lv2='" + sEc2 + "' ORDER BY tab_cod DESC";
                        DataTable t2 = _clsFun.FillTabSql("TabEcrLv3", s, true, _strConSql);
                        if (t2.Rows.Count > 0)
                            sEc3 = (Convert.ToInt16(t2.Rows[0]["tab_cod"]) + 1).ToString("000");

                        x = t2.NewRow();
                        x["tab_cod"] = sEc3;
                        x["tab_lv2"] = sEc2;
                        x["tab_lv1"] = sEc1;
                        x["tab_des"] = sLd3;
                        x["tab_ann"] = false;
                        s = _clsFun.SqlInsertRow("TabEcrLv3", t2, x);
                        _clsFun.SqlWrite(s, _strConSql);

                        x = t.NewRow();
                        x["tab_for"] = strFor;
                        x["tab_tip"] = "EC3";
                        x["tab_cof"] = sLc1 + sLc2 + sLc3;
                        x["tab_cod"] = sEc3;

                        s = _clsFun.SqlInsertRow("TabTabFornitori", t, x);
                        _clsFun.SqlWrite(s, _strConSql);
                    }
                    else
                        sEc3 = (string)t.Rows[0]["tab_cod"];
                }
            }

            return sEc1 +";"+ sEc2 +";"+ sEc3;
        }

        private string ForDivRegola(string strTip, string strReg)
        {
            string sRes = "";

            if (strReg.Contains(strTip))
            {
                string[] aReg = strReg.Trim().Split(',');

                foreach (string s1 in aReg)
                {
                    string[] a2 = s1.Split('-');
                    if (a2[0].Trim() == strTip)
                        sRes = a2[1];
                }
            }

            return sRes;
        }

        private string DivTraSeek(string strTra)
        {
            string sResTip = "";
            string sResLeg = "";
            string sResFil = "";
            string sRes = "";

            string s = "SELECT * FROM AnaForDivTestate WHERE trt_for='" + strTra + "' ORDER BY trt_tip";
            //DataTable tTrt = _clsFun.FillTabSql("AnaForDivTestate", s, false, _strConSql);
            DataTable tTrt = _clsFun.FillTabMdb("AnaForDivTestate", s, false, _strConSqlMdb);

            Boolean b = false;

            string sKey = "";                   //Per evitare record doppi con testate e righe

            foreach (DataRow y in tTrt.Rows)
            {
                if (sKey != (string)y["trt_tip"])
                {
                    sKey = (string)y["trt_tip"];

                    if ((string)y["trt_tip"] == "ART" || (string)y["trt_tip"] == "FAT" || (string)y["trt_tip"] == "OFF")
                    {
                        string sPth = ((string)y["trt_pth"]).Trim();

                        //string sPathFor = ((string)y["tab_pth"]).Trim();

                        s = ((string)y["trt_fil"]).Trim();
                        string[] aFil = s.Split(',');

                        string[] sFils = Directory.GetFiles(sPth);
                        foreach (string sFi in sFils)
                        {
                            string sFil = sFi;

                            foreach (string ss in aFil)
                            {
                                s = Path.GetFileName(sFil);

                                if (ss.Length > 0 && ss.Substring(0, 1) == "*")
                                {
                                    string sss = ss.Replace("*", "");

                                    if (s.Contains(sss))
                                    {
                                        b = true;
                                        sResTip = (string)y["trt_tip"];
                                        sResLeg = (string)y["trt_leg"];
                                        sResFil = sFi;

                                        sRes += sResTip + "+" + sResFil + "+" + sResLeg + "-";

                                    }
                                }
                                else if (ss.Length > 0 && ss.Contains("*"))
                                {
                                    s = Path.GetFileName(sFil);

                                    string[] a = ss.Split('*');

                                    if(s.Substring(0, a[0].Length) == a[0] && s.Substring(s.Length - a[1].Length) == a[1])
                                    {
                                        b = true;
                                        sResTip = (string)y["trt_tip"];
                                        sResLeg = (string)y["trt_leg"];
                                        sResFil = sFi;

                                        sRes += sResTip + "+" + sResFil + "+" + sResLeg + "-";

                                    }
                                }
                                else if (ss != "" && s.Length >= ss.Length && s.Substring(0, ss.Length).ToLower() == ss.ToLower())
                                {

                                    b = true;

                                    string sReg = "";
                                    if (!DBNull.Value.Equals(y["trt_reg"]) && ((string)y["trt_reg"]).Contains("FILE"))
                                    {
                                        string[] a = ((string)y["trt_reg"]).Split(',');
                                        foreach (string sx in a)
                                        {
                                            if (sx.Contains("FILE"))
                                            {
                                                string[] aa = sx.Trim().Split('-');
                                                if(aa.Length > 3)
                                                {
                                                    int iPos = Convert.ToInt16(aa[1]);
                                                    int iLen = Convert.ToInt16(aa[2]);
                                                    string sStr = aa[3];

                                                    s = Path.GetFileName(sFil);

                                                    if (s.Substring(iPos, iLen) != sStr)
                                                        b = false;
                                                }
                                            }
                                        }
                                    }

                                    if (b)
                                    {
                                        sResTip = (string)y["trt_tip"];
                                        sResLeg = (string)y["trt_leg"];
                                        sResFil = sFi;

                                        sRes += sResTip + "+" + sResFil + "+" + sResLeg + "-";
                                    }
                                }
                                else if (ss.Contains('*'))
                                {
                                    string[] af = ss.Split('*');

                                    s = Path.GetFileName(sFil).ToLower();

                                    if (s.Substring(0, af[0].Length) == af[0] && s.Substring(s.Length - af[1].Length, af[1].Length) == af[1])
                                    {
                                        b = true;
                                        sResTip = (string)y["trt_tip"];
                                        sResLeg = (string)y["trt_leg"];
                                        sResFil = sFi;

                                        sRes += sResTip + "+" + sResFil + "+" + sResLeg + "-";
                                    }
                                }


                            }
                            //if (b)
                            //break;
                        }
                    }
                }
                //if (b)
                    //break;
            }

            //s = "";
            //if(sResFil != "")
            //    s = sResTip + "+" + sResFil + "+" + sResLeg;

            return sRes;
        }

        private string DivTraRow(string strRig, DataTable tabTra, DataTable tabTes, DataTable tabRig, DataTable tabDie, string strPar, string strRes, string strEcr)
        {
            DataRow rowRow = tabRig.NewRow();
            rowRow["div_imp"] = 0;
            rowRow["div_qta"] = 0;
            rowRow["div_cos"] = 0;

            string s = "";
            string[] a;
            string sTmp = "";
            DataRow[] j;
            string sTipRec = "";                                    //Se testata scrivo in tabTes altrimenti in tabRig
            string sTrtCod = "";                                    //Codice tracciato corrispondente
            string sTipRig = "";

            a = strPar.Split('|');
            string sTva = a[0];
            string sFil = a[1];
            string sFor = a[2];
            string sFod = "";
            if (a.Length > 4)
                sFod = a[4];

            string sDivNum = "";                                    //NUmero divulgazione
            string sTesRig = "N";                                   //Tracciato con testate e righe
            Boolean bOk = true;

            foreach (DataRow y in tabTra.Rows)
            {
                Boolean b = true;
                sTipRec = "";

                if (((string)y["trt_reg"]).Trim() != "")             //Regola testate e righe
                {
                    s = ((string)y["trt_reg"]).Trim();
                    a = s.Split('-');

                    if (a.Length > 1 && a[0] == "FILTRO")
                    {
                        int iPos = Convert.ToInt16(a[1]);
                        int iLen = Convert.ToInt16(a[2]);
                        string sVal = a[3];

                        s = strRig.Substring(iPos, iLen);

                        if (s != sVal)
                            bOk = false;

                        sTrtCod = (string)y["trt_cod"];
                        b = false;
                        break;
                    }
                    else if (a.Length > 1 && a[0] == "TIPRIG")
                    {
                        int iPos = Convert.ToInt16(a[1]);
                        int iLen = Convert.ToInt16(a[2]);
                        string sVal = a[3];
                        sTipRig = strRig.Substring(iPos, iLen);
                        sTrtCod = (string)y["trt_cod"];
                        string[] aa = a[4].Split('=');          //Verifico la riga di defaul per una lettura normale
                        if (aa[1] == sTipRig)                   //Se diverso da DEFAULT filtro i campi
                            sTipRig = "";
                        else
                        {
                            aa = a[5].Split('=');
                            sTipRig = "2";
                            //sTipRec = aa[1];
                        }
                        bOk = true;
                        b = false;
                        break;
                    }
                    else if (a.Length == 2 && a[0] == "SEPARATORE")      //File con campi separati (";")
                    {
                        sTipRec = a[1].Trim();
                        sTrtCod = (string)y["trt_cod"];
                        b = false;
                        break;
                    }
                    else if (((string)y["trt_reg"]).Contains("SEPARATORE"))      //File con campi separati (";")
                    {
                        sTipRec = ForDivRegola("SEPARATORE", (string)y["trt_reg"]);

                        //sTipRec = s;

                        sTrtCod = (string)y["trt_cod"];
                        b = false;
                        break;
                    }
                    else if (a.Length > 3 && a[0] != "FILE")
                    {
                        b = false;
                        int iPos = Convert.ToInt16(a[0]);
                        int iLen = Convert.ToInt16(a[1]);
                        string sVal = a[2].Trim();
                        s = strRig.Substring(iPos, iLen);

                        if (s == sVal)
                        {
                            sTipRec = a[3].Trim();
                            sTrtCod = (string)y["trt_cod"];
                            break;
                        }
                    }

                }

                if(b)
                {
                    //Console.WriteLine("zzz");
                    string s2 = ((string)y["trt_fil"]).Trim().ToLower();

                    s = Path.GetFileName(sFil).ToLower();

                    Console.WriteLine("xxxx");

                    if ((string)y["trt_tip"] == sTva && s.Length >= s2.Length && s.Substring(0, s2.Length) == s2)
                    {
                        sTrtCod = (string)y["trt_cod"];
                        if (sTva == "EAN")
                            sTipRec = "Barcode";
                        break;
                    }
                }
            }

            s = "";

            if (bOk && sTrtCod != "")
            {
                foreach (DataRow y in tabTra.Rows)
                {
                    if ((string)y["trt_cod"] == sTrtCod)
                    {
                        Boolean b = false;
                        string sFld = (string)y["tra_fld"];
                        int iPos = Convert.ToInt16(y["tra_pos"]);
                        int iLen = Convert.ToInt16(y["tra_len"]);
                        string sReg = ((string)y["tra_reg"]).Trim();

                        if (sFld == "for_iva")
                            Console.WriteLine("zzzzzzzzzzzzzzz");

                        //if ((iPos >= 0 && iLen > 0 && strRig.Length >= iPos + iLen) || iLen == 999 || sReg.Contains("FILE"))

                        if (iPos >= 0 && iLen > 0 && strRig.Length >= iPos + iLen)
                            b = true;
                        if (iLen == 999)
                            b = true;
                        if (sReg.Contains("FILE"))
                            b = true;

                        if (sTipRig != "")
                            Console.WriteLine("xxxxxxxx");

                        if (b && sTipRig != "")
                        {
                            if (!sReg.Contains("TIPRIG"))
                                b = false;
                            else
                            {
                                string[] aa = sReg.Split(',');
                                foreach (string ss in aa)
                                {
                                    if (ss.Contains("TIPRIG"))
                                    {
                                        string[] aaa = ss.Split('=');
                                        if (!aaa[1].Contains(sTipRig))
                                            b = false;
                                    }
                                }
                            }
                        }

                        if (b && sTipRig != "")
                            Console.WriteLine("xxxxxxxx");

                        if (b)
                        {
                            string sVal = "";

                            if (sReg.Contains("FILE"))
                            {
                                string[] aReg = sReg.Split('-');

                                foreach (string sRg in aReg)
                                {
                                    Console.WriteLine("zzzzzz");

                                    string sSep = "";
                                    string sOri = "";

                                    if (sRg.Contains('+'))
                                    {
                                        a = sRg.Split('+');

                                        foreach (string ss in a)
                                        {
                                            string[] aa = ss.Split('=');

                                            if (aa.Length > 1)
                                            {
                                                if (aa[0] == "ORI")
                                                    sOri = aa[1];
                                                else if (aa[0] == "SEP")
                                                    sSep = aa[1];
                                            }
                                        }

                                        if (sOri == "FILE")
                                        {
                                            string sFi = Path.GetFileNameWithoutExtension(sFil);

                                            if (sSep == "UNDERSCORE")
                                            {
                                                a = sFi.Split('_');
                                                sVal = a[iPos];
                                            }
                                        }
                                    }
                                    else
                                        sReg = sRg;
                                }
                            }
                            else if (sTipRec == "PUNTOVIRGOLA" || sTipRec == "PUNTOESCLAMATIVO")
                            {
                                char[] cSep = new char[] { ';' };
                                if (sTipRec == "PUNTOESCLAMATIVO")
                                    cSep = new char[] { '!' };

                                string[] aRig = strRig.Split(cSep);
                                sVal = aRig[iPos - 1].Trim();

                                if (iLen == 999)
                                {
                                    sVal = "";

                                    for (int ii = iPos - 1; ii < aRig.Length; ii++)
                                    {
                                        s = aRig[ii].Trim();

                                        if (_clsFun.Numerico(s, "0123456789"))
                                        {
                                            s = Convert.ToDouble(s).ToString();

                                            if (s.Length > 13)
                                                s = "";
                                            else if (s.Length < 13)
                                                s = s.PadLeft(13, Convert.ToChar(' '));

                                            sVal += s;
                                        }
                                    }
                                }
                                else if (sFld == "div_ard" && sVal.Length > 50)
                                    sVal = sVal.Substring(0, 50);
                            }
                            else
                                sVal = strRig.Substring(iPos, iLen).Trim();

                            if (sVal != "")
                            {

                                //if (sFld == "div_ec1")
                                //    Console.WriteLine("zzzzzzzzz");

                                if (sReg.Length > 4 && sReg.Substring(0, 4) == "SBST")
                                {
                                    a = sReg.Substring(5).Split(',');

                                    Console.WriteLine("zzzzzzzzz");

                                    if(a.Length > 1)
                                    {
                                        int iIni = Convert.ToInt16(a[0]);
                                        int iFin = Convert.ToInt16(a[1]);

                                        rowRow[sFld] = sVal.Substring(iIni, iFin);
                                    }

                                }
                                else if (sReg.Length > 2 && sReg.Substring(0, 2) == "IF")
                                {
                                    if (sVal == "")
                                        Console.WriteLine("zzzz");

                                    s = sReg.Substring(3);
                                    a = s.Split('=');
                                    if (a.Length > 2)
                                    {
                                        //if (sVal == a[0])
                                        if (a[0].Contains(sVal))
                                            sVal = a[1];
                                        else
                                            sVal = a[2];
                                        rowRow[sFld] = sVal;

                                        sTmp += sFld + " " + sVal + " | ";
                                    }
                                    else if (a.Length > 1)
                                    {
                                        if (a[0] == "" || a[0].Contains(sVal))
                                            sVal = a[1];
                                        rowRow[sFld] = sVal;
                                        sTmp += sFld + " " + sVal + " | ";
                                    }
                                }
                                else if (sReg.Contains("AAAAMMDD"))
                                {
                                    rowRow[sFld] = _clsFun.Str2Day(sVal);
                                    sTmp += sFld + " " + sVal + " | ";
                                }
                                else if (sReg.Contains("AAMMDD"))
                                {
                                    rowRow[sFld] = _clsFun.Str2Day(sVal);
                                    sTmp += sFld + " " + sVal + " | ";
                                }
                                else if (sReg.Contains("DDMMAAAA"))
                                {
                                    sVal = sVal.Substring(4, 4) + sVal.Substring(2, 2) + sVal.Substring(0, 2);

                                    rowRow[sFld] = _clsFun.Str2Day(sVal);
                                    sTmp += sFld + " " + sVal + " | ";
                                }
                                else if (sReg.Contains("DD-MM-AAAA"))
                                {
                                    //sVal = sVal.Substring(4, 4) + sVal.Substring(2, 2) + sVal.Substring(0, 2);

                                    rowRow[sFld] = sVal;
                                    sTmp += sFld + " " + sVal + " | ";
                                }
                                else if (sReg == "INT")
                                {
                                    rowRow[sFld] = Convert.ToInt16(sVal);
                                    sTmp += sFld + " " + sVal + " | ";
                                }
                                else if (sReg.Length > 3 && sReg.Substring(0, 3) == "DEC")
                                {
                                    if (_clsFun.Numerico(sVal))
                                    {
                                        int dDec = Convert.ToInt16(sReg.Substring(4));
                                        decimal dVal = Convert.ToDecimal(sVal);
                                        if (dDec > 0 && dVal > 0)
                                        {
                                            rowRow[sFld] = dVal / dDec;
                                            sTmp += sFld + " " + sVal + " | ";
                                        }
                                    }
                                }
                                else if (sReg.Length > 5 && sReg.Substring(0, 4) == "PADL")
                                {
                                    if (_clsFun.Numerico(sVal))
                                    {
                                        string sFill = sReg.Substring(5);

                                        s = sFill + sVal.Trim();

                                        sVal = s.Substring(s.Length - sFill.Length, sFill.Length);
                                        rowRow[sFld] = sVal;
                                        sTmp += sFld + " " + sVal + " | ";
                                    }
                                }
                                else
                                {
                                    try
                                    {
                                        rowRow[sFld] = sVal;
                                        sTmp += sFld + " " + sVal + " | ";
                                    }
                                    catch (Exception ex)
                                    {
                                        Console.WriteLine(ex.Message);
                                        //this._clsMail.InvMail("edp@spacrovigo.org", "EDP", "paolo.secchettin@spacrovigo.org", "", "Errori su SpacServer", "clsGenFun.SqlWrite", "", "");
                                    }

                                }
                            }
                        }
                    }
                }

                string sNri = "";

                if (bOk)
                {
                    if (sTipRec == "Testata")
                    {
                        sDivNum = DivTraTestata(rowRow, tabTes, strPar);
                        sTesRig = "S";
                    }
                    else
                    {
                        sNri = "00001";

                        a = strRes.Split('-');
                        if (a.Length > 1)
                        {
                            sDivNum = a[0];
                            sNri = a[1];

                            if (sNri == "")
                                sNri = "0";

                            sNri = (Convert.ToInt32(sNri) + 1).ToString("00000");
                        }

                        if (sNri == "00016")
                            Console.WriteLine("xxxx");

                        if (a.Length > 2)
                            sTesRig = a[2];

                        rowRow["div_tva"] = sTva;
                        rowRow["div_num"] = sDivNum;
                        rowRow["div_nri"] = sNri;

                        if (sTva == "FAT" && (DBNull.Value.Equals(rowRow["div_imp"]) || (decimal)rowRow["div_imp"] == 0))
                        {
                            if ((decimal)rowRow["div_qta"] != 0 && (decimal)rowRow["div_cos"] > 0)
                                rowRow["div_imp"] = (decimal)rowRow["div_qta"] * (decimal)rowRow["div_cos"];
                        }

                        if (DBNull.Value.Equals(rowRow["div_var"]))
                            rowRow["div_var"] = "";
                        if (DBNull.Value.Equals(rowRow["div_dva"]))
                            rowRow["div_dva"] = DateTime.Today;
                        if (DBNull.Value.Equals(rowRow["div_prp"]))
                            rowRow["div_prp"] = 0;
                        if (DBNull.Value.Equals(rowRow["div_eti"]))
                            rowRow["div_eti"] = "";
                        if (DBNull.Value.Equals(rowRow["div_ddo"]))
                            rowRow["div_ddo"] = DateTime.Today;
                        if (DBNull.Value.Equals(rowRow["div_ndo"]))
                            rowRow["div_ndo"] = "";

                        if(strEcr.Length > 10)
                        {
                            a = strEcr.Split(';');

                            if(a.Length > 2)
                            {
                                rowRow["div_ec1"] = a[0];
                                rowRow["div_ec2"] = a[1];
                                rowRow["div_ec3"] = a[2];
                            }
                        }

                        //if (sTesRig != "S" ||  sNri == "00001")
                        if (sTesRig != "S")
                            sDivNum = DivTraTestata(rowRow, tabTes, strPar);

                        string sEan = "";

                        if (sTipRec != "Barcode")
                        {
                            rowRow["div_num"] = sDivNum;

                            if (!DBNull.Value.Equals(rowRow["div_ean"]) && (string)rowRow["div_ean"] != "")
                            {
                                sEan = (string)rowRow["div_ean"];
                                rowRow["div_ean"] = "";
                            }
                            if (sTipRig == "")
                                tabRig.Rows.Add(rowRow);
                        }

                        if (sEan != "")
                        {
                            Console.WriteLine("zzzzzzzz");

                            for (int ii = 0; ii < sEan.Length; ii += 13)
                            {
                                if (sEan.Substring(0, 1) == "0")
                                    Console.WriteLine("xxxxxxx");
                                if (sEan.Length > 13)
                                    Console.WriteLine("xxxxxxx");

                                string sBar = "";

                                if (ii + 13 > sEan.Length)
                                    sBar = sEan.Substring(ii);
                                else
                                {
                                    if (sEan.Length > 13 && sEan.Length < 16)   //Scarto barcode > 14 e < 16 perchè significa che non ci sono 2 barcode sulla stessa riga
                                        sBar = "";
                                    else
                                        sBar = sEan.Substring(ii, 13);
                                }
                                if (_clsFun.Numerico(sBar, "0123456789"))
                                {
                                    sBar = Convert.ToDouble(sBar).ToString();

                                    Boolean bEan = false;

                                    if (sBar.Trim() != "" && sBar.Length > 7 && _clsFun.Numerico(sBar, "0123456789"))
                                        bEan = true;
                                    else if(sFod == "APSHOP")
                                        bEan = true;

                                    if(bEan)
                                    {
                                        s = sBar;
                                        if(sBar.Length > 7)
                                            s = sBar.Substring(0, sBar.Length - 1) + new clsCtrlCodici().FindMod10Digit(sBar.Substring(0, sBar.Length - 1));

                                        j = tabDie.Select("die_arf='" + rowRow["div_arf"] + "' AND die_ean='" + s + "'");
                                        if (j.Length == 0)
                                        {
                                            if (s == "0178")
                                                Console.WriteLine("aaaaaaaaaa");

                                            DataRow yy = tabDie.NewRow();
                                            yy["die_num"] = sDivNum;
                                            yy["die_for"] = sFor;
                                            yy["die_arf"] = rowRow["div_arf"];
                                            yy["die_ean"] = s;

                                            yy["die_bil"] = false;
                                            if (!DBNull.Value.Equals(rowRow["div_bil"]) && (string)rowRow["div_bil"] == "1")
                                                yy["die_bil"] = true;
                                            tabDie.Rows.Add(yy);

                                            if (s != sBar)
                                                _clsFun.ErrorLog("Divulgazione " + sFor, "Corretto check digit " + sBar + " -> " + s);
                                        }
                                    }

                                }
                            }

                            s = "xxxxxxxxxxxxxxxxxxxxx";
                        }
                    }

                    strRes = sDivNum + "-" + sNri + "-" + sTesRig;
                }
            }
            return strRes;
        }

        private string DivTraTestata(DataRow rowRow, DataTable tabTes, string strPar)
        {
            string[] a = strPar.Split('|');
            string sTva = a[0];
            string sFil = a[1];
            string sFor = a[2];
            string sFod = a[4];

            DataRow[] j;

            string sNdo = "";
            if (!DBNull.Value.Equals(rowRow["div_ndo"]))
                sNdo = (string)rowRow["div_ndo"];

            DateTime dDdo = DateTime.Today;
            if (!DBNull.Value.Equals(rowRow["div_ddo"]))
                dDdo = (DateTime)rowRow["div_ddo"];

            if (sTva == "EAN")
                sTva = "ART";

            if (sTva == "OFF")
            {
                if (DBNull.Value.Equals(rowRow["div_off"]) || ((string)rowRow["div_off"]).Trim() == "")
                    sNdo = "OFF";
                else
                    sNdo = (string)rowRow["div_off"];
            }

            string sDivNum = "";

            j = tabTes.Select("dit_tva='" + sTva + "' AND dit_ndo='" + sNdo + "' AND dit_ddo=#" + dDdo.ToString("MM/dd/yyyy") + "#");
            if (j.Length == 0)
            {
                sDivNum = DitNewNum();

                DataRow x = tabTes.NewRow();
                x["dit_cho"] = "E";
                x["dit_ddo"] = dDdo;

                //x["dit_des"] = "Variazioni anagrafiche";
                //if (sTva == "FAT")
                //    x["dit_des"] = "Fattura";
                x["dit_des"] = sFod;

                x["dit_fil"] = Path.GetFileName(sFil);
                x["dit_for"] = sFor;
                x["dit_ndo"] = sNdo;
                x["dit_neg"] = "";
                x["dit_nri"] = 1;
                x["dit_num"] = sDivNum;
                x["dit_nva"] = "";
                x["dit_pth"] = Path.GetDirectoryName(sFil);
                //x["dit_sif"] = "";
                //x["dit_sii"] = "";
                //x["dit_sof"] = "";
                //x["dit_soi"] = "";

                if (sTva == "OFF")
                {
                    if (DBNull.Value.Equals(rowRow["off_soi"])) // || ((string)rowRow["off_soi"]).Trim() == "")
                    {
                        x["dit_soi"] = DateTime.Today;
                        x["dit_sof"] = DateTime.Today;
                    }
                    else
                    {
                        x["dit_soi"] = (DateTime)rowRow["off_soi"];
                        x["dit_sof"] = (DateTime)rowRow["off_soo"];
                    }
                }

                x["dit_tfo"] = "";
                x["dit_tva"] = sTva;
                tabTes.Rows.Add(x);
            }
            else
            {
                sDivNum = (string)j[0]["dit_num"];
            }

            return sDivNum;
        }

        private void  DivFor2Anag(string strFor, string strFod)
        {
            string s = "SELECT * FROM AnaFornitori WHERE for_cod='" + strFor + "'";
            DataTable t = _clsFun.FillTabSql("AnaFornitori", s, true, _strConSql);
            if(t.Rows.Count == 0)
            {
                DataRow x = t.NewRow();
                x["for_cod"] = strFor;
                x["for_des"] = strFod;

                s = _clsFun.SqlInsertRow("AnaFornitori", t, x);
                _clsFun.SqlWrite(s, _strConSql);
            }
        }

        private void CtrlTabMarchio(string strMar, string strMad)
        {
            ArrayList aWhe = new ArrayList();
            aWhe.Add("tab_cod");
            ArrayList aExl = new ArrayList();

            string s = "SELECT * FROM TabMarchio WHERE tab_cod='" + strMar + "'";
            DataTable t = _clsFun.FillTabSql("TabMarchio", s, true, _strConSql);
            DataRow x = t.NewRow();
            x["tab_cod"] = strMar;
            x["tab_des"] = strMad;

            if (t.Rows.Count == 0)
                s = _clsFun.SqlInsertRow("TabMarchio", t, x);
            else
                s = _clsFun.SqlUpdRowIdx("TabMarchio", t, t.Rows[0], x, aExl);
            if (s != "")
                _clsFun.SqlWrite(s, _strConSql);
        }

    }

}
