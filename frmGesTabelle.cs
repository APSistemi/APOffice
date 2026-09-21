using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;

namespace APOffice
{
    public partial class frmGesTabelle : Form
    {
        clsDefine _clsDef = new clsDefine();
        clsFuncs _clsFun = new clsFuncs();
        clsQuery _clsQry = new clsQuery();

        public string _strTab = "";
        public string _strTabDes = "";
        public Boolean _bolCodAlf = false;
        public Boolean _bolColAnn = false;
        public int _intCodLen = 3;
        public int _intDesLen = 20;
        public Boolean _bolSql = false;

        //private string _strConMdb = "";
        private string _strConSql = "";

        private BindingSource _mBs1 = new BindingSource();
        private DataSet _mDs1 = new DataSet();
        //private OleDbDataAdapter _mDa1 = new OleDbDataAdapter();
        private SqlDataAdapter _sDa1 = new SqlDataAdapter();
        private DataTable _mTb1 = new DataTable();

        public frmGesTabelle()
        {
            InitializeComponent();
            new clsGesGraph().SetGraph(this, 0);;
        }

        private void frmGesTabelle_Load(object sender, EventArgs e)
        {
            this.Text = _strTabDes;

            //_strConMdb = _clsFun.ConMdb("");
            _strConSql = _clsFun.ConSql("");

            SetDgv1();
        }

        private void dgv1_KeyDown(object sender, KeyEventArgs e)
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
            dgv1.CommitEdit(DataGridViewDataErrorContexts.Commit);

            if ((DataTable)_mBs1.DataSource != null)
            {
                dgv1.Refresh();

                DataTable tTmp = _mTb1.GetChanges();

                _mBs1.EndEdit();
                //_mDa1.Update((DataTable)_mBs1.DataSource);

                //if (_bolSql)
                _sDa1.Update((DataTable)_mBs1.DataSource);
                //else
                //    _mDa1.Update((DataTable)_mBs1.DataSource);

                string s = _clsFun.ParGet(clsDefine.enuParametri.Par016PathDivNegozi, _strConSql);
                if(tTmp != null && tTmp.Rows.Count > 0 && s != "" && _strTab != "TabParametri")
                    new clsVariazioni().DivNegTabelle(s, tTmp);
            }

            this.Close();
        }

        private void SetDgv1()
        {
            dgv1.AutoGenerateColumns = false;
            dgv1.AllowUserToAddRows = false;
            //dgv1.VirtualMode = true;
            //dgv1.Dock = DockStyle.Fill;

            DataGridViewTextBoxColumn cTbc;
            DataGridViewCheckBoxColumn cCbc;
            DataGridViewComboBoxColumn cCmb;
            DataGridViewButtonColumn cBtn;

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "tab_cod";
            cTbc.Name = "Codice";
            cTbc.Width = 50;
            cTbc.ValueType = typeof(string);
            cTbc.MaxInputLength = _intCodLen;
            dgv1.Columns.Add(cTbc);
            dgv1.Columns["Codice"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.ValueType = typeof(string);
            cTbc.DataPropertyName = "tab_des";
            cTbc.Name = "Descrizione";
            cTbc.Width = 250;
            cTbc.MaxInputLength = _intDesLen;
            dgv1.Columns.Add(cTbc);

            if (_strTab == "TabNote")
                dgv1.Columns["Descrizione"].Visible = false;

            if (_strTab == "TabUtenti")
            {
                cTbc = new DataGridViewTextBoxColumn();
                cTbc.ValueType = typeof(string);
                cTbc.DataPropertyName = "tab_pwd";
                cTbc.Name = "Pwd";
                cTbc.Width = 60;
                cTbc.MaxInputLength = 10;
                dgv1.Columns.Add(cTbc);

                cTbc = new DataGridViewTextBoxColumn();
                cTbc.ValueType = typeof(string);
                cTbc.DataPropertyName = "tab_liv";
                cTbc.Name = "Livello";
                cTbc.Width = 60;
                cTbc.MaxInputLength = 2;
                dgv1.Columns.Add(cTbc);
            }
            else if (_strTab == "TabNumeratori")
            {
                cTbc = new DataGridViewTextBoxColumn();
                cTbc.ValueType = typeof(string);
                cTbc.DataPropertyName = "tab_yea";
                cTbc.Name = "Anno";
                cTbc.Width = 60;
                cTbc.MaxInputLength = 4;
                dgv1.Columns.Add(cTbc);
                dgv1.Columns[cTbc.Name].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                dgv1.Columns[cTbc.Name].DefaultCellStyle.Format = "####";

                cTbc = new DataGridViewTextBoxColumn();
                cTbc.ValueType = typeof(string);
                cTbc.DataPropertyName = "tab_val";
                cTbc.Name = "Numero";
                cTbc.Width = 60;
                cTbc.MaxInputLength = 10;
                dgv1.Columns.Add(cTbc);
                dgv1.Columns[cTbc.Name].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                dgv1.Columns[cTbc.Name].DefaultCellStyle.Format = "###,##0";
            }
            else if (_strTab == "TabParametri")
            {
                cTbc = new DataGridViewTextBoxColumn();
                cTbc.ValueType = typeof(string);
                cTbc.DataPropertyName = "tab_val";
                cTbc.Name = "Parametro";
                cTbc.Width = 200;
                cTbc.MaxInputLength = 100;
                dgv1.Columns.Add(cTbc);
                //dgv1.Columns[cTbc.Name].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                //dgv1.Columns[cTbc.Name].DefaultCellStyle.Format = "###,##0";
            }
            else if (_strTab == "TabTerm")
            {
                cTbc = new DataGridViewTextBoxColumn();
                cTbc.ValueType = typeof(string);
                cTbc.DataPropertyName = "tab_bat";
                cTbc.Name = "Batch invio aggiornamento";
                cTbc.Width = 100;
                cTbc.MaxInputLength = 50;
                dgv1.Columns.Add(cTbc);
                //dgv1.Columns[cTbc.Name].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                //dgv1.Columns[cTbc.Name].DefaultCellStyle.Format = "###,##0";

                cTbc = new DataGridViewTextBoxColumn();
                cTbc.ValueType = typeof(string);
                cTbc.DataPropertyName = "tab_bao";
                cTbc.Name = "Batch ric ordini";
                cTbc.Width = 100;
                cTbc.MaxInputLength = 50;
                dgv1.Columns.Add(cTbc);
                //dgv1.Columns[cTbc.Name].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                //dgv1.Columns[cTbc.Name].DefaultCellStyle.Format = "###,##0";

                cTbc = new DataGridViewTextBoxColumn();
                cTbc.ValueType = typeof(string);
                cTbc.DataPropertyName = "tab_div";
                cTbc.Name = "File divulgaz";
                cTbc.Width = 100;
                cTbc.MaxInputLength = 50;
                dgv1.Columns.Add(cTbc);
                //dgv1.Columns[cTbc.Name].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                //dgv1.Columns[cTbc.Name].DefaultCellStyle.Format = "###,##0";

                cTbc = new DataGridViewTextBoxColumn();
                cTbc.ValueType = typeof(string);
                cTbc.DataPropertyName = "tab_hst";
                cTbc.Name = "Ftp host";
                cTbc.Width = 100;
                cTbc.MaxInputLength = 50;
                cTbc.ToolTipText = "Host ftp";
                dgv1.Columns.Add(cTbc);

                cTbc = new DataGridViewTextBoxColumn();
                cTbc.ValueType = typeof(string);
                cTbc.DataPropertyName = "tab_usr";
                cTbc.Name = "Ftp user";
                cTbc.Width = 100;
                cTbc.MaxInputLength = 20;
                cTbc.ToolTipText = "User ftp";
                dgv1.Columns.Add(cTbc);

                cTbc = new DataGridViewTextBoxColumn();
                cTbc.ValueType = typeof(string);
                cTbc.DataPropertyName = "tab_pwd";
                cTbc.Name = "Ftp pwd";
                cTbc.Width = 100;
                cTbc.MaxInputLength = 20;
                cTbc.ToolTipText = "Pwd ftp";
                dgv1.Columns.Add(cTbc);

                cTbc = new DataGridViewTextBoxColumn();
                cTbc.ValueType = typeof(string);
                cTbc.DataPropertyName = "tab_fpt";
                cTbc.Name = "Ftp percorso";
                cTbc.Width = 100;
                cTbc.MaxInputLength = 100;
                cTbc.ToolTipText = "Percorso ftp ; codici terminali";
                dgv1.Columns.Add(cTbc);

                cCbc = new DataGridViewCheckBoxColumn();
                cCbc.ValueType = typeof(Boolean);
                cCbc.DataPropertyName = "tab_var";
                cCbc.Name = "Variazioni";
                cCbc.Width = 50;
                dgv1.Columns.Add(cCbc);
            }
            else if (_strTab == "TabForImport")
            {
                //string s = Path.GetDirectoryName(_clsDef.FILEINI) + "\\ApForTracciati.mdb";
                string s = _clsFun.ConMdb(Path.GetDirectoryName(_clsDef.FILEINI) + "\\ApForTracciati.mdb");
                //if (!File.Exists(s))
                //{
                //    MessageBox.Show("File tracciati mancante!", "CONTROLLO PARAMETRI", MessageBoxButtons.OK, MessageBoxIcon.Error);
                //    Esci();
                //}
                //else
                //{
                //DataTable tTra = _clsFun.FillTabSql("AnaForDivTipi", "SELECT * FROM AnaForDivTipi WHERE trk_sta='A'", false, _strConSql);
                DataTable tTra = _clsFun.FillTabMdb("AnaForDivTipi", "SELECT * FROM AnaForDivTipi WHERE trk_sta='A'", false, s);
                DataRow x = tTra.NewRow();
                x["trk_cod"] = "";
                x["trk_des"] = "  Non definito";
                tTra.Rows.InsertAt(x, 0);

                cTbc = new DataGridViewTextBoxColumn();
                cTbc.ValueType = typeof(string);
                cTbc.DataPropertyName = "tab_tip";
                cTbc.Name = "Tipo divulgazione";
                cTbc.Width = 100;
                cTbc.MaxInputLength = 10;
                cTbc.ToolTipText = "Sigla identificativa del fornitore";
                dgv1.Columns.Add(cTbc);

                cTbc = new DataGridViewTextBoxColumn();
                cTbc.ValueType = typeof(string);
                cTbc.DataPropertyName = "tab_neg";
                cTbc.Name = "Codice negozio";
                cTbc.Width = 100;
                cTbc.MaxInputLength = 20;
                cTbc.ToolTipText = "Codice negozio per il fornitore";
                dgv1.Columns.Add(cTbc);

                cTbc = new DataGridViewTextBoxColumn();
                cTbc.ValueType = typeof(string);
                cTbc.DataPropertyName = "tab_pth";
                cTbc.Name = "Percorso";
                cTbc.Width = 100;
                cTbc.MaxInputLength = 50;
                cTbc.ToolTipText = "Path lettura file";
                dgv1.Columns.Add(cTbc);

                cTbc = new DataGridViewTextBoxColumn();
                cTbc.ValueType = typeof(string);
                cTbc.DataPropertyName = "tab_fil";
                cTbc.Name = "Variazioni";
                cTbc.Width = 100;
                cTbc.MaxInputLength = 100;
                cTbc.ToolTipText = "Prefisso nome file variazioni";
                dgv1.Columns.Add(cTbc);

                cTbc = new DataGridViewTextBoxColumn();
                cTbc.ValueType = typeof(string);
                cTbc.DataPropertyName = "tab_fio";
                cTbc.Name = "Offerte";
                cTbc.Width = 100;
                cTbc.MaxInputLength = 50;
                cTbc.ToolTipText = "Prefisso nome file Offerte";
                dgv1.Columns.Add(cTbc);

                cTbc = new DataGridViewTextBoxColumn();
                cTbc.ValueType = typeof(string);
                cTbc.DataPropertyName = "tab_fat";
                cTbc.Name = "Fatture";
                cTbc.Width = 100;
                cTbc.MaxInputLength = 50;
                cTbc.ToolTipText = "Prefisso nome file Fatture";
                dgv1.Columns.Add(cTbc);

                cTbc = new DataGridViewTextBoxColumn();
                cTbc.ValueType = typeof(string);
                cTbc.DataPropertyName = "tab_ord";
                cTbc.Name = "Ordini";
                cTbc.Width = 100;
                cTbc.MaxInputLength = 50;
                cTbc.ToolTipText = "Prefisso nome file Fatture";
                dgv1.Columns.Add(cTbc);

                cTbc = new DataGridViewTextBoxColumn();
                cTbc.ValueType = typeof(string);
                cTbc.DataPropertyName = "tab_hst";
                cTbc.Name = "Ftp host";
                cTbc.Width = 100;
                cTbc.MaxInputLength = 120;
                cTbc.ToolTipText = "Host ftp";
                dgv1.Columns.Add(cTbc);

                cTbc = new DataGridViewTextBoxColumn();
                cTbc.ValueType = typeof(string);
                cTbc.DataPropertyName = "tab_usr";
                cTbc.Name = "Ftp user";
                cTbc.Width = 100;
                cTbc.MaxInputLength = 20;
                cTbc.ToolTipText = "User ftp";
                dgv1.Columns.Add(cTbc);

                cTbc = new DataGridViewTextBoxColumn();
                cTbc.ValueType = typeof(string);
                cTbc.DataPropertyName = "tab_pwd";
                cTbc.Name = "Ftp pwd";
                cTbc.Width = 100;
                cTbc.MaxInputLength = 20;
                cTbc.ToolTipText = "Pwd ftp";
                dgv1.Columns.Add(cTbc);

                cTbc = new DataGridViewTextBoxColumn();
                cTbc.ValueType = typeof(string);
                cTbc.DataPropertyName = "tab_fpt";
                cTbc.Name = "Ftp percorso";
                cTbc.Width = 100;
                cTbc.MaxInputLength = 50;
                cTbc.ToolTipText = "Percorso ftp";
                dgv1.Columns.Add(cTbc);

                cTbc = new DataGridViewTextBoxColumn();
                cTbc.ValueType = typeof(string);
                cTbc.DataPropertyName = "tab_uso";
                cTbc.Name = "Ftp user ordini";
                cTbc.Width = 100;
                cTbc.MaxInputLength = 30;
                cTbc.ToolTipText = "User ftp";
                dgv1.Columns.Add(cTbc);

                cTbc = new DataGridViewTextBoxColumn();
                cTbc.ValueType = typeof(string);
                cTbc.DataPropertyName = "tab_pwo";
                cTbc.Name = "Ftp pwd ordini";
                cTbc.Width = 100;
                cTbc.MaxInputLength = 30;
                cTbc.ToolTipText = "Pwd ftp";
                dgv1.Columns.Add(cTbc);

                cTbc = new DataGridViewTextBoxColumn();
                cTbc.ValueType = typeof(string);
                cTbc.DataPropertyName = "tab_mai";
                cTbc.Name = "Ordini via mail";
                cTbc.Width = 200;
                cTbc.MaxInputLength = 100;
                cTbc.ToolTipText = "Pwd ftp";
                dgv1.Columns.Add(cTbc);

                cCbc = new DataGridViewCheckBoxColumn();
                cCbc.ValueType = typeof(Boolean);
                cCbc.DataPropertyName = "tab_nor";
                cCbc.Name = "Escluso da ordini";
                cCbc.Width = 30;
                dgv1.Columns.Add(cCbc);

                cCmb = new DataGridViewComboBoxColumn();
                cCmb.DataPropertyName = "tab_tra";
                cCmb.Name = "Tracciato";
                cCmb.Width = 200;
                cCmb.DataSource = tTra;
                cCmb.ValueMember = "trk_cod";
                cCmb.DisplayMember = "trk_des";
                dgv1.Columns.Add((DataGridViewColumn)cCmb);
            }
            else if (_strTab == "TabTipoGrammatura")
            {
                cTbc = new DataGridViewTextBoxColumn();
                cTbc.ValueType = typeof(string);
                cTbc.DataPropertyName = "tab_val";
                cTbc.Name = "Conversione";
                cTbc.Width = 80;
                cTbc.MaxInputLength = 20;
                dgv1.Columns.Add(cTbc);
                //dgv1.Columns[cTbc.Name].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                //dgv1.Columns[cTbc.Name].DefaultCellStyle.Format = "###,##0";
            }
            else if (_strTab == "TabIva")
            {
                cTbc = new DataGridViewTextBoxColumn();
                cTbc.ValueType = typeof(decimal);
                cTbc.DataPropertyName = "tab_ali";
                cTbc.Name = "Aliquota";
                cTbc.Width = 60;
                cTbc.MaxInputLength = 6;
                dgv1.Columns.Add(cTbc);
                dgv1.Columns[cTbc.Name].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                dgv1.Columns[cTbc.Name].DefaultCellStyle.Format = "#0";

                //DataTable tCnf = _clsQry.ConfSeek("", "POS");
                //if ((string)tCnf.Rows[0]["cnf_pos"] == Convert.ToInt16((object)clsDefine.enuPos.posNcr745x).ToString("00"))
                //{
                cTbc = new DataGridViewTextBoxColumn();
                cTbc.ValueType = typeof(string);
                cTbc.DataPropertyName = "tab_pos";
                cTbc.Name = "Cassa";
                cTbc.Width = 30;
                cTbc.MaxInputLength = 3;
                dgv1.Columns.Add(cTbc);

                cTbc = new DataGridViewTextBoxColumn();
                cTbc.ValueType = typeof(string);
                cTbc.DataPropertyName = "tab_ele";
                cTbc.Name = "fattura elettronica";
                cTbc.Width = 30;
                cTbc.MaxInputLength = 6;
                cTbc.ToolTipText = "Natura per codice esente (N1 art.15, N2.2 giornali)";
                dgv1.Columns.Add(cTbc);
                //}
            }
            else if (_strTab == "TabGrammatura")
            {
                cTbc = new DataGridViewTextBoxColumn();
                cTbc.ValueType = typeof(decimal);
                cTbc.DataPropertyName = "tab_val";
                cTbc.Name = "Conversione in KG";
                cTbc.Width = 60;
                cTbc.MaxInputLength = 6;
                dgv1.Columns.Add(cTbc);
                dgv1.Columns[cTbc.Name].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                dgv1.Columns[cTbc.Name].DefaultCellStyle.Format = "#0";
            }
            else if (_strTab == "TabNote")
            {
                cTbc = new DataGridViewTextBoxColumn();
                cTbc.ValueType = typeof(decimal);
                cTbc.DataPropertyName = "tab_row";
                cTbc.Name = "Riga";
                cTbc.Width = 60;
                cTbc.MaxInputLength = 2;
                dgv1.Columns.Add(cTbc);

                cTbc = new DataGridViewTextBoxColumn();
                cTbc.ValueType = typeof(decimal);
                cTbc.DataPropertyName = "tab_txt";
                cTbc.Name = "Testo";
                cTbc.Width = 360;
                cTbc.MaxInputLength = 200;
                dgv1.Columns.Add(cTbc);
            }
            else if (_strTab == "TabListini")
            {
                cTbc = new DataGridViewTextBoxColumn();
                cTbc.ValueType = typeof(decimal);
                cTbc.DataPropertyName = "tab_tip";
                cTbc.Name = "Tipo";
                cTbc.Width = 30;
                cTbc.MaxInputLength = 1;
                dgv1.Columns.Add(cTbc);

                cTbc = new DataGridViewTextBoxColumn();
                cTbc.ValueType = typeof(decimal);
                cTbc.DataPropertyName = "tab_iva";
                cTbc.Name = "IVA";
                cTbc.Width = 30;
                cTbc.MaxInputLength = 1;
                cTbc.ToolTipText = "'S' IVA compresa";
                dgv1.Columns.Add(cTbc);
            }
            else if (_strTab == "TabNegozi")
            {
                cTbc = new DataGridViewTextBoxColumn();
                cTbc.ValueType = typeof(string);
                cTbc.DataPropertyName = "tab_lis";
                cTbc.Name = "Listino";
                cTbc.Width = 50;
                cTbc.MaxInputLength = 3;
                dgv1.Columns.Add(cTbc);

                cTbc = new DataGridViewTextBoxColumn();
                cTbc.ValueType = typeof(string);
                cTbc.DataPropertyName = "tab_tip";
                cTbc.Name = "Tipo";
                cTbc.Width = 40;
                cTbc.MaxInputLength = 1;
                cTbc.ToolTipText = "'L'ocale, 'E'sterno";
                dgv1.Columns.Add(cTbc);

                cTbc = new DataGridViewTextBoxColumn();
                cTbc.ValueType = typeof(string);
                cTbc.DataPropertyName = "tab_pos";
                cTbc.Name = "Casse";
                cTbc.Width = 100;
                cTbc.MaxInputLength = 30;
                cTbc.ToolTipText = "'1,2,3,...";
                dgv1.Columns.Add(cTbc);

                cTbc = new DataGridViewTextBoxColumn();
                cTbc.ValueType = typeof(string);
                cTbc.DataPropertyName = "tab_rep";
                cTbc.Name = "Reparti";
                cTbc.Width = 100;
                cTbc.MaxInputLength = 50;
                cTbc.ToolTipText = "'I1,E10,...(I=incluso E=escluso)";
                dgv1.Columns.Add(cTbc);

                cTbc = new DataGridViewTextBoxColumn();
                cTbc.ValueType = typeof(string);
                cTbc.DataPropertyName = "tab_mag";
                cTbc.Name = "Magazzino";
                cTbc.Width = 50;
                cTbc.MaxInputLength = 2;
                cTbc.ToolTipText = "Codice magazzino";
                dgv1.Columns.Add(cTbc);
            }
            else if (_strTab == "TabStato")
            {
                cCbc = new DataGridViewCheckBoxColumn();
                cCbc.ValueType = typeof(Boolean);
                cCbc.DataPropertyName = "tab_art";
                cCbc.Name = "Articoli";
                cCbc.Width = 30;
                dgv1.Columns.Add(cCbc);

                cCbc = new DataGridViewCheckBoxColumn();
                cCbc.ValueType = typeof(Boolean);
                cCbc.DataPropertyName = "tab_off";
                cCbc.Name = "Offerte";
                cCbc.Width = 30;
                dgv1.Columns.Add(cCbc);

                cCbc = new DataGridViewCheckBoxColumn();
                cCbc.ValueType = typeof(Boolean);
                cCbc.DataPropertyName = "tab_doc";
                cCbc.Name = "Documenti";
                cCbc.Width = 30;
                dgv1.Columns.Add(cCbc);

                cCbc = new DataGridViewCheckBoxColumn();
                cCbc.ValueType = typeof(Boolean);
                cCbc.DataPropertyName = "tab_lot";
                cCbc.Name = "Lotti";
                cCbc.Width = 30;
                dgv1.Columns.Add(cCbc);
            }
            else if (_strTab == "TabStatoDivulgazioni")
            {
                cCbc = new DataGridViewCheckBoxColumn();
                cCbc.ValueType = typeof(Boolean);
                cCbc.DataPropertyName = "tab_off";
                cCbc.Name = "Offerte";
                cCbc.Width = 50;
                dgv1.Columns.Add(cCbc);

                cCbc = new DataGridViewCheckBoxColumn();
                cCbc.ValueType = typeof(string);
                cCbc.DataPropertyName = "tab_var";
                cCbc.Name = "Variazioni";
                cCbc.Width = 50;
                dgv1.Columns.Add(cCbc);

            }
            else if (_strTab == "TabPos")
            {
                cTbc = new DataGridViewTextBoxColumn();
                cTbc.ValueType = typeof(string);
                cTbc.DataPropertyName = "tab_var";
                cTbc.Name = "Path variazioni";
                cTbc.Width = 130;
                cTbc.MaxInputLength = 50;
                dgv1.Columns.Add(cTbc);

                cTbc = new DataGridViewTextBoxColumn();
                cTbc.ValueType = typeof(string);
                cTbc.DataPropertyName = "tab_vao";
                cTbc.Name = "Path offerte";
                cTbc.Width = 130;
                cTbc.MaxInputLength = 50;
                dgv1.Columns.Add(cTbc);

                cTbc = new DataGridViewTextBoxColumn();
                cTbc.ValueType = typeof(string);
                cTbc.DataPropertyName = "tab_ven";
                cTbc.Name = "Path venduto";
                cTbc.Width = 130;
                cTbc.MaxInputLength = 50;
                dgv1.Columns.Add(cTbc);

                cTbc = new DataGridViewTextBoxColumn();
                cTbc.ValueType = typeof(string);
                cTbc.DataPropertyName = "tab_mat";
                cTbc.Name = "Matricole";
                cTbc.Width = 230;
                cTbc.MaxInputLength = 60;
                dgv1.Columns.Add(cTbc);
            }
            else if (_strTab == "TabBilance")
            {
                cTbc = new DataGridViewTextBoxColumn();
                cTbc.ValueType = typeof(string);
                cTbc.DataPropertyName = "tab_var";
                cTbc.Name = "Path variazioni";
                cTbc.Width = 150;
                cTbc.MaxInputLength = 100;
                dgv1.Columns.Add(cTbc);
            }
            else if (_strTab == "TabReparti")
            {
                cTbc = new DataGridViewTextBoxColumn();
                cTbc.ValueType = typeof(string);
                cTbc.DataPropertyName = "tab_bil";
                cTbc.Name = "Reparto bilance";
                cTbc.Width = 50;
                cTbc.MaxInputLength = 4;
                dgv1.Columns.Add(cTbc);
            }
            else if (_strTab == "TabRepBilance")
            {
                cTbc = new DataGridViewTextBoxColumn();
                cTbc.ValueType = typeof(string);
                cTbc.DataPropertyName = "tab_tar";
                cTbc.Name = "Tara gr";
                cTbc.Width = 30;
                cTbc.MaxInputLength = 3;
                dgv1.Columns.Add(cTbc);
                dgv1.Columns[cTbc.Name].DefaultCellStyle.Format = "###";

                cTbc = new DataGridViewTextBoxColumn();
                cTbc.ValueType = typeof(string);
                cTbc.DataPropertyName = "tab_gru";
                cTbc.Name = "Gruppo";
                cTbc.Width = 30;
                cTbc.MaxInputLength = 2;
                dgv1.Columns.Add(cTbc);
                dgv1.Columns[cTbc.Name].DefaultCellStyle.Format = "##";
            }
            else if (_strTab == "TabMovCausali")
            {
                cTbc = new DataGridViewTextBoxColumn();
                cTbc.ValueType = typeof(string);
                cTbc.DataPropertyName = "tab_cfo";
                cTbc.Name = "CLI/FOR";
                cTbc.Width = 60;
                cTbc.MaxInputLength = 3;
                dgv1.Columns.Add(cTbc);

                cTbc = new DataGridViewTextBoxColumn();
                cTbc.ValueType = typeof(string);
                cTbc.DataPropertyName = "tab_mag";
                cTbc.Name = "Magazzino";
                cTbc.Width = 60;
                cTbc.MaxInputLength = 2;
                dgv1.Columns.Add(cTbc);
                dgv1.Columns[cTbc.Name].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                dgv1.Columns[cTbc.Name].ToolTipText = "01/02/...";
                dgv1.Columns[cTbc.Name].DefaultCellStyle.Format = "00";

                cTbc = new DataGridViewTextBoxColumn();
                cTbc.ValueType = typeof(string);
                cTbc.DataPropertyName = "tab_sgm";
                cTbc.Name = "Segno";
                cTbc.Width = 60;
                cTbc.MaxInputLength = 1;
                dgv1.Columns.Add(cTbc);
                dgv1.Columns[cTbc.Name].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                dgv1.Columns[cTbc.Name].ToolTipText = "+/-/ ";

                cTbc = new DataGridViewTextBoxColumn();
                cTbc.ValueType = typeof(string);
                cTbc.DataPropertyName = "tab_tip";
                cTbc.Name = "Tipo";
                cTbc.Width = 60;
                cTbc.MaxInputLength = 10;
                dgv1.Columns.Add(cTbc);
                dgv1.Columns[cTbc.Name].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                dgv1.Columns[cTbc.Name].ToolTipText = "P = proforma (non genera numero doc.) K Peso";

                //cCbc = new DataGridViewCheckBoxColumn();
                //cCbc.ValueType = typeof(Boolean);
                //cCbc.DataPropertyName = "tab_pos";
                //cCbc.Name = "In cassa";
                //cCbc.Width = 50;
                //dgv1.Columns.Add(cCbc);

                cTbc = new DataGridViewTextBoxColumn();
                cTbc.ValueType = typeof(string);
                cTbc.DataPropertyName = "tab_pos";
                cTbc.Name = "In cassa";
                cTbc.Width = 70;
                cTbc.MaxInputLength = 3;
                dgv1.Columns.Add(cTbc);
                dgv1.Columns[cTbc.Name].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                dgv1.Columns[cTbc.Name].ToolTipText = "FIS=Stampante fiscale, PRN=stampante";

                cTbc = new DataGridViewTextBoxColumn();
                cTbc.ValueType = typeof(string);
                cTbc.DataPropertyName = "tab_num";
                cTbc.Name = "Numeratore";
                cTbc.Width = 70;
                cTbc.MaxInputLength = 3;
                dgv1.Columns.Add(cTbc);
                dgv1.Columns[cTbc.Name].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                dgv1.Columns[cTbc.Name].ToolTipText = "Codice tabella numeratori";
            }
            else if (_strTab == "TabFidCampagne")
            {
                DataTable tGru = _clsFun.FillTabSql("TabFidGruppi", "SELECT * FROM TabFidGruppi", false, _strConSql);
                DataRow x = tGru.NewRow();
                x["tab_cod"] = "";
                x["tab_des"] = "Non definito";
                tGru.Rows.Add(x);

                cCmb = new DataGridViewComboBoxColumn();
                cCmb.DataPropertyName = "tab_gru";
                cCmb.Name = "Gruppo";
                cCmb.Width = 150;
                cCmb.DataSource = tGru;
                cCmb.ValueMember = "tab_cod";
                cCmb.DisplayMember = "tab_des";
                dgv1.Columns.Add((DataGridViewColumn)cCmb);

                cTbc = new DataGridViewTextBoxColumn();
                cTbc.ValueType = typeof(DateTime);
                cTbc.DataPropertyName = "tab_dti";
                cTbc.Name = "dal";
                cTbc.Width = 80;
                cTbc.MaxInputLength = 10;
                dgv1.Columns.Add(cTbc);
                //dgv1.Columns[cTbc.Name].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                dgv1.Columns[cTbc.Name].DefaultCellStyle.Format = "dd/MM/yyyy";

                cTbc = new DataGridViewTextBoxColumn();
                cTbc.ValueType = typeof(DateTime);
                cTbc.DataPropertyName = "tab_dtf";
                cTbc.Name = "al";
                cTbc.Width = 80;
                cTbc.MaxInputLength = 10;
                dgv1.Columns.Add(cTbc);
                //dgv1.Columns[cTbc.Name].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                dgv1.Columns[cTbc.Name].DefaultCellStyle.Format = "dd/MM/yyyy";
            }
            else if (_strTab == "TabProvince")
            {
                cTbc = new DataGridViewTextBoxColumn();
                cTbc.ValueType = typeof(string);
                cTbc.DataPropertyName = "tab_reg";
                cTbc.Name = "Regione";
                cTbc.Width = 150;
                cTbc.MaxInputLength = 50;
                //cTbc.ToolTipText = "Codice negozio";
                dgv1.Columns.Add(cTbc);
            }
            else if (_strTab == "TabFidGruppi")
            {
                DataTable tCam = _clsFun.FillTabSql("TabFidCampagne", "SELECT * FROM TabFidCampagne", false, _strConSql);
                DataRow x = tCam.NewRow();
                x["tab_cod"] = "";
                x["tab_des"] = "Non definito";
                tCam.Rows.Add(x);

                cCmb = new DataGridViewComboBoxColumn();
                cCmb.DataPropertyName = "tab_cam";
                cCmb.Name = "Campagna";
                cCmb.Width = 200;
                cCmb.DataSource = tCam;
                cCmb.ValueMember = "tab_cod";
                cCmb.DisplayMember = "tab_des";
                dgv1.Columns.Add((DataGridViewColumn)cCmb);

                cTbc = new DataGridViewTextBoxColumn();
                cTbc.ValueType = typeof(string);
                cTbc.DataPropertyName = "tab_neg";
                cTbc.Name = "Codice negozio";
                cTbc.Width = 70;
                cTbc.MaxInputLength = 20;
                //cTbc.ToolTipText = "Codice negozio";
                dgv1.Columns.Add(cTbc);
            }
            else if (_strTab == "TabCauCassa")
            {
                cTbc = new DataGridViewTextBoxColumn();
                cTbc.ValueType = typeof(string);
                cTbc.DataPropertyName = "tab_tri";
                cTbc.Name = "Tipo riga";
                cTbc.Width = 30;
                cTbc.MaxInputLength = 1;
                cTbc.ToolTipText = "1-pagamenti, 2-sconti, 3-punti, 4-cassa";
                dgv1.Columns.Add(cTbc);

                cTbc = new DataGridViewTextBoxColumn();
                cTbc.ValueType = typeof(string);
                cTbc.DataPropertyName = "tab_tip";
                cTbc.Name = "Tipo operazione";
                cTbc.Width = 40;
                cTbc.MaxInputLength = 3;
                cTbc.ToolTipText = "PAG-pagamento, SCP-sconto %, SCV-sconto a valore, SCT-sconto fidelity, PUN-punti";
                dgv1.Columns.Add(cTbc);

                cTbc = new DataGridViewTextBoxColumn();
                cTbc.ValueType = typeof(string);
                cTbc.DataPropertyName = "tab_ord";
                cTbc.Name = "Ordinamento in visualizzazione";
                cTbc.Width = 30;
                cTbc.MaxInputLength = 1;
                dgv1.Columns.Add(cTbc);

                cTbc = new DataGridViewTextBoxColumn();
                cTbc.ValueType = typeof(string);
                cTbc.DataPropertyName = "tab_sgn";
                cTbc.Name = "Segno";
                cTbc.Width = 30;
                cTbc.MaxInputLength = 1;
                cTbc.ToolTipText = "+/-/ ";
                dgv1.Columns.Add(cTbc);

                cTbc = new DataGridViewTextBoxColumn();
                cTbc.ValueType = typeof(string);
                cTbc.DataPropertyName = "tab_ppo";
                cTbc.Name = "Pos";
                cTbc.Width = 30;
                cTbc.MaxInputLength = 1;
                cTbc.ToolTipText = "1-contanti, 2-bancomat CCR, 3-assegno";
                dgv1.Columns.Add(cTbc);

                cTbc = new DataGridViewTextBoxColumn();
                cTbc.ValueType = typeof(string);
                cTbc.DataPropertyName = "tab_key";
                cTbc.Name = "Key";
                cTbc.Width = 40;
                cTbc.MaxInputLength = 3;
                cTbc.ToolTipText = "CON-contanti, RES-resto, PUN-punti +, PUM-punti -, PRE-premio";
                dgv1.Columns.Add(cTbc);
            }
            else if (_strTab == "TabDocumenti")
            {
                cTbc = new DataGridViewTextBoxColumn();
                cTbc.ValueType = typeof(string);
                cTbc.DataPropertyName = "tab_suf";
                cTbc.Name = "Suffisso";
                cTbc.Width = 80;
                cTbc.MaxInputLength = 10;
                cTbc.ToolTipText = "Sigla dopo il numero in stampa del documento";
                dgv1.Columns.Add(cTbc);
            }
            else if (_strTab == "TabDocTipo")
            {
                cTbc = new DataGridViewTextBoxColumn();
                cTbc.ValueType = typeof(string);
                cTbc.DataPropertyName = "tab_cfo";
                cTbc.Name = "Cl/fo";
                cTbc.Width = 50;
                cTbc.MaxInputLength = 3;
                dgv1.Columns.Add(cTbc);

                cTbc = new DataGridViewTextBoxColumn();
                cTbc.ValueType = typeof(string);
                cTbc.DataPropertyName = "tab_tip";
                cTbc.Name = "Tipo";
                cTbc.Width = 50;
                cTbc.MaxInputLength = 1;
                cTbc.ToolTipText = "K-peso";
                dgv1.Columns.Add(cTbc);

                //cCbc = new DataGridViewCheckBoxColumn();
                //cCbc.ValueType = typeof(Boolean);
                //cCbc.DataPropertyName = "tab_pos";
                //cCbc.Name = "In cassa";
                //cCbc.Width = 50;
                //dgv1.Columns.Add(cCbc);

                cTbc = new DataGridViewTextBoxColumn();
                cTbc.ValueType = typeof(string);
                cTbc.DataPropertyName = "tab_pos";
                cTbc.Name = "In cassa";
                cTbc.Width = 70;
                cTbc.MaxInputLength = 3;
                dgv1.Columns.Add(cTbc);
                dgv1.Columns[cTbc.Name].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                dgv1.Columns[cTbc.Name].ToolTipText = "FIS=Stampante fiscale, PRN=stampante";
            }
            else if (_strTab == "TabDocTpd")
            {
                cTbc = new DataGridViewTextBoxColumn();
                cTbc.ValueType = typeof(string);
                cTbc.DataPropertyName = "tab_num";
                cTbc.Name = "Numeratore";
                cTbc.Width = 60;
                cTbc.MaxInputLength = 3;
                cTbc.ToolTipText = "Numeratore progressivo da tabella numeratori";
                cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                dgv1.Columns.Add(cTbc);

                cTbc = new DataGridViewTextBoxColumn();
                cTbc.ValueType = typeof(string);
                cTbc.DataPropertyName = "tab_ele";
                cTbc.Name = "Fatt.elettronica";
                cTbc.Width = 60;
                cTbc.MaxInputLength = 4;
                cTbc.ToolTipText = "Tipo documento in fattura elettronica (alternativo a TD01)";
                cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                dgv1.Columns.Add(cTbc);
            }
            else if (_strTab == "TabEcrLv2")
            {
                //if (_clsFun.ParGet(clsDefine.enuParametri.ParGesTagColori, _strConSql) == "S")
                //{
                //    DataTable tPet = _clsFun.FillTabSql("TabPettini", "SELECT * FROM TabPettini", false, _strConSql);
                //    DataRow x = tPet.NewRow();
                //    x["tab_cod"] = "";
                //    x["tab_des"] = "Non definito";
                //    tPet.Rows.Add(x);

                //    cCmb = new DataGridViewComboBoxColumn();
                //    cCmb.DataPropertyName = "tab_lv1";
                //    cCmb.Name = "Pettine";
                //    cCmb.Width = 120;
                //    cCmb.DataSource = tPet;
                //    cCmb.ValueMember = "tab_cod";
                //    cCmb.DisplayMember = "tab_des";
                //    dgv1.Columns.Add((DataGridViewColumn)cCmb);
                //}
            }
            else if (_strTab == "TabEtichette")
            {
                cTbc = new DataGridViewTextBoxColumn();
                cTbc.ValueType = typeof(string);
                cTbc.DataPropertyName = "tab_msg";
                cTbc.Name = "Messaggio (separazione con ';')";
                cTbc.Width = 300;
                cTbc.MaxInputLength = 200;
                dgv1.Columns.Add(cTbc);

                cTbc = new DataGridViewTextBoxColumn();
                cTbc.ValueType = typeof(string);
                cTbc.DataPropertyName = "tab_par";
                cTbc.Name = "Parametri";
                cTbc.Width = 100;
                cTbc.MaxInputLength = 50;
                dgv1.Columns.Add(cTbc);

                DataGridViewComboBoxColumn cCmbSta = new DataGridViewComboBoxColumn();
                DataTable tStaEti = new DataTable();
                tStaEti.Columns.Add("cod", typeof(string));
                tStaEti.Columns.Add("des", typeof(string));
                tStaEti.Rows.Add("S", "S - Stampata (con richiesta)");
                tStaEti.Rows.Add("N", "N - Da Stampare");
                tStaEti.Rows.Add("E", "E - Cancellata");
                cCmbSta.DataSource = tStaEti;
                cCmbSta.ValueMember = "cod";
                cCmbSta.DisplayMember = "des";
                cCmbSta.DataPropertyName = "tab_sta";
                cCmbSta.Name = "Post-Stampa (S/N/E)";
                cCmbSta.Width = 175;
                dgv1.Columns.Add(cCmbSta);
            }
            else if (_strTab == "TabEquPrezzi")
            {
                cBtn = new DataGridViewButtonColumn();
                cBtn.UseColumnTextForButtonValue = true;
                cBtn.FlatStyle = FlatStyle.System;
                cBtn.Name = "Articoli";
                cBtn.Text = "...";
                cBtn.Width = 70;
                cBtn.ValueType = typeof(string);
                cBtn.ReadOnly = false;
                cBtn.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                dgv1.Columns.Add(cBtn);
            }
            else if (_strTab == "TabPagamenti")
            {
                DataTable tFtp = _clsFun.FillTabSql("TabPagETipoPagamento", "SELECT * FROM TabPagETipoPagamento ORDER BY tab_cod", false, _strConSql);
                DataRow x = tFtp.NewRow();
                x["tab_cod"] = "";
                x["tab_des"] = "Non definito";
                tFtp.Rows.Add(x);

                DataTable tFte = _clsFun.FillTabSql("TabPagEModalita", "SELECT * FROM TabPagEModalita ORDER BY tab_cod", false, _strConSql);
                x = tFte.NewRow();
                x["tab_cod"] = "";
                x["tab_des"] = "Non definito";
                tFte.Rows.Add(x);

                DataTable tPdt = _clsFun.FillTabSql("TabPagDate", "SELECT * FROM TabPagDate ORDER BY tab_cod", false, _strConSql);
                x = tPdt.NewRow();
                x["tab_cod"] = "";
                x["tab_des"] = "Non definito";
                tPdt.Rows.Add(x);

                //cTbc = new DataGridViewTextBoxColumn();
                //cTbc.ValueType = typeof(string);
                //cTbc.DataPropertyName = "tab_ftp";
                //cTbc.Name = "Condizioni pagamento fattura elettronica";
                //cTbc.Width = 100;
                //cTbc.MaxInputLength = 4;
                //dgv1.Columns.Add(cTbc);

                cCmb = new DataGridViewComboBoxColumn();
                cCmb.DataPropertyName = "tab_ftp";
                cCmb.Name = "Condizioni pagamento fattura elettronica";
                cCmb.Width = 150;
                cCmb.DataSource = tFtp;
                cCmb.ValueMember = "tab_cod";
                cCmb.DisplayMember = "tab_des";
                dgv1.Columns.Add((DataGridViewColumn)cCmb);

                //cTbc = new DataGridViewTextBoxColumn();
                //cTbc.ValueType = typeof(string);
                //cTbc.DataPropertyName = "tab_fte";
                //cTbc.Name = "Modalità pagamento fattura elettronica";
                //cTbc.Width = 100;
                //cTbc.MaxInputLength = 4;
                //dgv1.Columns.Add(cTbc);

                cCmb = new DataGridViewComboBoxColumn();
                cCmb.DataPropertyName = "tab_fte";
                cCmb.Name = "Modalità pagamento fattura elettronica";
                cCmb.Width = 150;
                cCmb.DataSource = tFte;
                cCmb.ValueMember = "tab_cod";
                cCmb.DisplayMember = "tab_des";
                dgv1.Columns.Add((DataGridViewColumn)cCmb);

                cCmb = new DataGridViewComboBoxColumn();
                cCmb.DataPropertyName = "tab_dpa";
                cCmb.Name = "Date pagamento";
                cCmb.Width = 150;
                cCmb.DataSource = tPdt;
                cCmb.ValueMember = "tab_cod";
                cCmb.DisplayMember = "tab_des";
                dgv1.Columns.Add((DataGridViewColumn)cCmb);
            }            

            if (_bolColAnn)
            {
                cCbc = new DataGridViewCheckBoxColumn();
                cCbc.ValueType = typeof(Boolean);
                cCbc.DataPropertyName = "tab_ann";
                cCbc.Name = "Ann.";
                cCbc.Width = 50;
                dgv1.Columns.Add(cCbc);
            }
            //if (_bolSql)
            FillDataSql();
            //else
            //    FillDataMdb();
        }
        
        //private void FillDataMdb()
        //{
        //    OleDbConnection cn = new OleDbConnection(_strConMdb);

        //    string s = "SELECT * FROM " + _strTab + " ORDER BY tab_cod";

        //    _mDa1 = new OleDbDataAdapter(s, cn);

        //    OleDbCommandBuilder cb = new OleDbCommandBuilder(_mDa1);

        //    _mDa1.Fill(_mTb1);

        //    _mBs1.DataSource = _mTb1;

        //    dgv1.Rows.Clear();
        //    dgv1.DataSource = _mBs1;
        //    dgv1.Refresh();
        //}

        private void FillDataSql()
        {
            SqlConnection cn = new SqlConnection(_strConSql);

            string s = "SELECT * FROM " + _strTab + " ORDER BY tab_cod";

            _sDa1 = new SqlDataAdapter(s, cn);

            SqlCommandBuilder cb = new SqlCommandBuilder(_sDa1);

            _mTb1.TableName = _strTab;

            _sDa1.Fill(_mTb1);

            _mBs1.DataSource = _mTb1;

            dgv1.Rows.Clear();
            dgv1.DataSource = _mBs1;
            dgv1.Refresh();
        }

        private void nuovoToolStripMenuItem_Click(object sender, EventArgs e)
        {
            dgv1.CommitEdit(DataGridViewDataErrorContexts.Commit);

            DataRow y = _clsFun.RowIniz(_mTb1);

            string sCod = "";

            if (_bolCodAlf)
                sCod = "";
            else
            {
                if (_mTb1.Rows.Count > 0)
                {
                    DataView v = new DataView(_mTb1, "", "tab_cod DESC", DataViewRowState.CurrentRows);
                    sCod = Convert.ToString(Convert.ToInt32(v[0]["tab_cod"]) + 1).PadLeft(_intCodLen, Convert.ToChar("0"));
                }
                else
                    sCod = "1".PadLeft(_intCodLen, Convert.ToChar("0"));
            }
            y["tab_cod"] = sCod;
            _mTb1.Rows.Add(y);
        }

        private void dgv1_CellValidating(object sender, DataGridViewCellValidatingEventArgs e)
        {
            dgv1.Rows[e.RowIndex].ErrorText = "";

            if (dgv1.Columns[e.ColumnIndex].DataPropertyName == "tab_cod")
            {
                string s = (string)e.FormattedValue;

                if (s.Length < _intCodLen)
                {
                    e.Cancel = true;
                    dgv1.Rows[e.RowIndex].ErrorText = "I dati inseriti non validi!";
                }
            }
        }

        //private void dgv1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        //{
        //    if(_strTab == "TabEquPrezzi" && e.ColumnIndex == 2)
        //    {
        //        CurrencyManager cm = dgv1.BindingContext[dgv1.DataSource, dgv1.DataMember] as CurrencyManager;
        //        if (cm.Position >= 0)
        //        {
        //            DataRowView r = cm.List[dgv1.CurrentRow.Index] as DataRowView;
        //            DataRow x = r.Row;
        //            string s = (string)x["tab_cod"];
        //            AnaEqp(s);
        //        }
        //    }
        //}

        //private void xxxdgv1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        //{

        //}

        private void AnaEqp(string strCod)
        {
            decimal d = 0;

            //DataRow[] j = ((DataView)dgv1.DataSource).Table.Select("LivIdx='" + _clsDef.LISPOS + "'");
            //if (j.Length > 0)
            //    d = (decimal)j[0]["LivPrv"];

            string sEqp = strCod;      // cmbArtEqp.SelectedValue.ToString();
            frmAnaArtEquivalenze f = new frmAnaArtEquivalenze();
            f._strTip = "SoloVideo";
            f._strEqp = sEqp;
            f._decPrv = d;
            f.ShowDialog();
        }

        private void dgv1_CellContentDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            if (_strTab == "TabEquPrezzi" && e.ColumnIndex == 2)
            {
                CurrencyManager cm = dgv1.BindingContext[dgv1.DataSource, dgv1.DataMember] as CurrencyManager;
                if (cm.Position >= 0)
                {
                    DataRowView r = cm.List[dgv1.CurrentRow.Index] as DataRowView;
                    DataRow x = r.Row;
                    string s = (string)x["tab_cod"];

                    AnaEqp(s);
                }
            }
        }

    }
}
