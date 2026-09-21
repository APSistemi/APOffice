using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace APOffice
{
    public partial class frmGesDocRaggruppa : Form
    {
        clsDefine _clsDef = new clsDefine();
        clsFuncs _clsFun = new clsFuncs();

        private const string TABANACLI = "AnaClienti";
        //private const string TABTABDOC = "TabDocumenti";
        private const string TABTABDTP = "TabDocTipo";
        private const string TABTABCAU = "TabMovCausali";

        private const string TABGESMOT = "GesMovTestate";
        private const string TABGESMOV = "GesMovimenti";
        private const string TABGESFAT = "GesFatTestate";

        private string _strConSql = "";
        private string _strSuf = "";

        public frmGesDocRaggruppa()
        {
            InitializeComponent();
            new clsGesGraph().SetGraph(this, 0);;
        }

        private void frmGesDocRaggruppa_Load(object sender, EventArgs e)
        {
            _strConSql = _clsFun.ConSql("");

            SetDgv1();
            FillTab();
        }
        private void frmGesDocRaggruppa_KeyDown(object sender, KeyEventArgs e)
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
            this.Close();
        }

        private void SetDgv1()
        {
            dgv1.AutoGenerateColumns = false;
            //dgv1.VirtualMode = true;
            //dgv1.Dock = DockStyle.Fill;
            dgv1.AllowUserToAddRows = false;
            dgv1.ReadOnly = false;
            dgv1.AllowUserToDeleteRows = false;
            //dgv1.DisplayedRowCount() = true;

            DataGridViewTextBoxColumn cTbc;
            DataGridViewCheckBoxColumn cCbc;
            //DataGridViewComboBoxColumn cCmb;

            //s += "fat_yfa AS DocYea, ";
            //s += "fat_nfa AS DocNum, ";
            //s += "fat_ndo AS DocNdo, ";
            //s += "fat_ddo AS DocDdo, ";
            //s += "fat_cfo AS DocCfo, ";
            //s += "AnaClienti.cli_des AS DocCfd, ";
            //s += "fat_ann AS DocAnn ";

            cCbc = new DataGridViewCheckBoxColumn();
            cCbc.DataPropertyName = "MotExl";
            cCbc.Name = "Escluso";
            cCbc.Width = 40;
            cCbc.ValueType = typeof(string);
            cCbc.ToolTipText = "Escluso dalla fatturazione";
            //cCbc.ReadOnly = false;
            dgv1.Columns.Add(cCbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "mot_ymo";
            cTbc.Name = "Anno";
            cTbc.Width = 50;
            cTbc.ValueType = typeof(string);
            cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            cTbc.ReadOnly = true;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "mot_nmo";
            cTbc.Name = "Numero";
            cTbc.Width = 50;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = true;
            //cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "MotCfd";
            cTbc.Name = "Cliente";
            cTbc.Width = 170;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = true;
            //cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "mot_ndo";
            cTbc.Name = "Documento";
            cTbc.Width = 80;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = true;
            //cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "mot_ddo";
            cTbc.Name = "Data";
            cTbc.Width = 90;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = true;
            //cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dgv1.Columns.Add(cTbc);

            cCbc = new DataGridViewCheckBoxColumn();
            cCbc.DataPropertyName = "mot_ann";
            cCbc.Name = "Ann.";
            cCbc.Width = 40;
            cCbc.ValueType = typeof(string);
            cTbc.ReadOnly = true;
            //cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dgv1.Columns.Add(cCbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "mot_nfa";
            cTbc.Name = "Mov.fattura";
            cTbc.Width = 80;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = true;
            //cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "MotMsg";
            cTbc.Name = "Messaggi";
            cTbc.Width = 200;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = true;
            //cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            cTbc.DefaultCellStyle.ForeColor = Color.Red;
            dgv1.Columns.Add(cTbc);
        }

        private void FillTab()
        {
            DataRow x;
            string s = "";
            DataTable t = new DataTable();

            s = "SELECT cli_cod, cli_des FROM AnaClienti ORDER BY cli_des";
            t = _clsFun.FillTabSql(TABANACLI, s, false, _strConSql);
            x = t.NewRow();
            x["cli_cod"] = "";
            x["cli_des"] = "  Non definito";
            t.Rows.InsertAt(x, 0);
            cmbFatCli.DataSource = t;
            cmbFatCli.DisplayMember = "cli_des";
            cmbFatCli.ValueMember = "cli_cod";
            cmbFatCli.SelectedValue = "";

            s = "SELECT * FROM TabMovCausali";
            t = _clsFun.FillTabSql(TABTABCAU, s, false, _strConSql);
            //x = t.NewRow();
            //x["tab_cod"] = "";
            //x["tab_des"] = "  Non definito";
            //t.Rows.InsertAt(x, 0);
            cmbMotCau.DataSource = t;
            cmbMotCau.DisplayMember = "tab_des";
            cmbMotCau.ValueMember = "tab_cod";
            cmbMotCau.SelectedValue = "001";
            //cmbMotCau.Enabled = false;

            s = "SELECT * FROM TabDocTipo WHERE Tab_cfo = '" + _clsDef.TIPCLI + "'";
            t = _clsFun.FillTabSql(TABTABDTP, s, false, _strConSql);
            //x = t.NewRow();
            //x["tab_cod"] = "";
            //x["tab_des"] = "  Non definito";
            //t.Rows.InsertAt(x, 0);
            cmbFatTpd.DataSource = t;
            cmbFatTpd.DisplayMember = "tab_des";
            cmbFatTpd.ValueMember = "tab_cod";
            cmbFatTpd.SelectedValue = "FD";
            cmbFatTpd.Enabled = true;

            if(t.Rows.Count == 0)
                MessageBox.Show("Tabella tipo documento non configurata!", "CONTROLLO TIPO DOCUMENTO", MessageBoxButtons.OK, MessageBoxIcon.Error);

            //s = "SELECT * FROM TabDocumenti WHERE tab_doc='F'";
            s = "SELECT * FROM TabDocumenti WHERE tab_cod='F'";
            t = _clsFun.FillTabSql("TabDocumenti", s, false, _strConSql);
            if (t.Rows.Count > 0 && !DBNull.Value.Equals(t.Rows[0]["tab_suf"]))
            {
                 s = (string)t.Rows[0]["tab_suf"];

                string[] a = s.Split(',');
                if (a.Length > 0)
                    _strSuf = a[0];

            }
        }

        private void btnSeekCli_Click(object sender, EventArgs e)
        {
            frmSeekCli f = new frmSeekCli();
            f._bolAnagra = true;
            f._bolSelect = true;
            f.ShowDialog();
            if (f._strCod != "")
            {
                if (f._tabTmp.Rows.Count > 0)
                {
                    cmbFatCli.SelectedValue = f._tabTmp.Rows[0]["tmp_cli"];

                    if (cmbFatCli.SelectedValue == null)
                    {
                        cmbFatCli.SelectedValue = f._tabTmp.Rows[0]["tmp_cli"];
                    }
                }
            }

        }

        private void btnEstrai_Click(object sender, EventArgs e)
        {
            if (cmbFatCli.SelectedValue.ToString() == null || cmbFatCli.SelectedValue.ToString() == "")
                MessageBox.Show("Cliente non selezionato!");
            else if (cmbMotCau.SelectedValue.ToString() == null || cmbMotCau.SelectedValue.ToString() == "")
                MessageBox.Show("Causale non selezionata!");
            else
                FillDati();
        }

        private void FillDati()
        {
            string s = "";

            s = "SELECT ";
            s += "mot_ymo, ";
            s += "mot_nmo, ";
            s += "mot_ndo, ";
            s += "mot_ddo, ";
            s += "mot_cfo, ";
            s += "AnaClienti.cli_des AS MotCfd, ";
            s += "mot_nfa, ";
            s += "mot_ann ";
            s += "FROM GesMovTestate ";
            s += "LEFT OUTER JOIN AnaClienti ON GesMovTestate.mot_cfo = AnaClienti.cli_cod ";
            s += "WHERE ";
            s += "mot_ddo<=" + _clsFun.DaySql( dtpMotFin.Value) + " AND ";

            if (cmbMotCau.SelectedValue.ToString() != "")
                s += "mot_cau='" + cmbMotCau.SelectedValue.ToString() + "' AND ";

            s += "mot_cfo ='" + cmbFatCli.SelectedValue.ToString() + "' AND ";
            s += "mot_nfa ='" + _clsDef.COD06X + "' AND ";
            s += "mot_ann =0 ";
            s += "ORDER BY mot_yfa, mot_nfa";

            DataTable t = _clsFun.FillTabSql(TABGESMOT, s, false, _strConSql);

            DataColumn c = new DataColumn();
            c.DataType = Type.GetType("System.Boolean");
            c.ColumnName = "MotExl";
            c.Caption = "Escluso";
            c.ReadOnly = false;
            c.Unique = false;
            c.DefaultValue = false;
            t.Columns.Add(c);

            c = new DataColumn();
            c.DataType = Type.GetType("System.String");
            c.ColumnName = "MotMsg";
            c.Caption = "Descrizione";
            c.MaxLength = 50;
            c.ReadOnly = false;
            c.DefaultValue = (String)"";
            t.Columns.Add(c);

            foreach(DataRow y in t.Rows)
            {
                if (((string)y["mot_ndo"]).Trim() == "")
                    y["MotMsg"] = "Numero documento non definito";
            }

            t.DefaultView.AllowDelete = false;
            t.DefaultView.AllowNew = false;
            dgv1.DataSource = t;
        }

        private void btnExl_Click(object sender, EventArgs e)
        {
            DataTable t = (DataTable)this.dgv1.DataSource;
            if (t != null)
            {
                foreach (DataRow y in t.Rows)
                {
                    if (!(bool)y["MotExl"])
                        y["MotExl"] = true;
                    else
                        y["MotExl"] = false;
                }
            }
            else
                MessageBox.Show("Non ci sono elementi estratti!");
        }

        private void btnOk_Click(object sender, EventArgs e)
        {
            Boolean b = true;

            if (dgv1.DataSource == null)
            {
                MessageBox.Show("Non ci sono documenti selezionate", "GENERAZIONE FATTURE", MessageBoxButtons.OK, MessageBoxIcon.Information); 
                b = false;
            }
            else
            {
                DataView v = new DataView((DataTable)dgv1.DataSource, "MotMsg<>''", "", DataViewRowState.CurrentRows);
                if (v.Count > 0)
                {
                    if (MessageBox.Show("Ci sono messaggi sulle righe documenti, procedi comunque?", "GENERAZIONE FATTURE", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
                        b = false;
                }
            }
            if(b)
                Mov2Fat();
        }

        private void Mov2Fat()
        {
            string sUbi = _clsFun.FileIni("R", clsDefine.enuIni.Ini13Ubicazione, "");
            if(sUbi == "")
                sUbi = _clsFun.FileIni("R", clsDefine.enuIni.Ini09CodiceAzienda, "");

            DataTable t = (DataTable)dgv1.DataSource;
            DataTable tTmp = new DataTable();
            int i = 0;
            string s = "";
            DataRow[] j;
            DataRow x;
            Boolean b = true;

            b = false;
            if (t == null)
                MessageBox.Show("Nessun documento selezionato!");
            else
            {
                DataView v = new DataView(t, "MotExl=0 AND mot_ann=0", "", DataViewRowState.CurrentRows);
                if (v.Count == 0)
                    MessageBox.Show("Nessun documento selezionato!");
                else
                    b = true;
            }

            if (cmbFatTpd.DataSource == null || cmbFatTpd.SelectedValue == null || cmbFatTpd.SelectedValue.ToString() == "")
            {
                MessageBox.Show("Tipo documento da stampare non selezionato!");
                b = false;
            }

            if(b)
            {
                i = 0;

                string sCli = cmbFatCli.SelectedValue.ToString(); 
                string sYea = dtpFatDdo.Value.Year.ToString();
                string sNmo = "";
                if (lblMotNfa.Text != "")
                    sNmo = lblMotNfa.Text;
                else
                    sNmo = _clsFun.NewNum(dtpMotFin.Value.Year.ToString(), clsDefine.enuNumeratori.NumGesMovFatture, 6, _strConSql);


                if (lblMotNfa.Text != "")
                    s = "SELECT * FROM GesFatTestate WHERE fat_nfa='" + sNmo + "'";
                else
                    s = "SELECT * FROM GesFatTestate WHERE fat_nfa='" + _clsDef.COD06X + "'";
                DataTable tFat = _clsFun.FillTabSql(TABGESFAT, s, true, _strConSql);

                x = tFat.NewRow();
                x["fat_ubi"] = sUbi;
                x["fat_yfa"] = sYea;
                x["fat_nfa"] = sNmo;
                x["fat_tpd"] = cmbFatTpd.SelectedValue.ToString();
                x["fat_ndo"] = "";
                x["fat_ddo"] = dtpFatDdo.Value;
                x["fat_no1"] = "";
                x["fat_tpg"] = "";
                x["fat_ann"] = 0;
                x["fat_cfo"] = sCli;
                x["fat_pvi"] = true;

                if (tFat.Rows.Count > 0)
                {
                    x = tFat.Rows[0];

                    s = _clsFun.SqlInsertRow(TABGESFAT, tFat, x);
                    _clsFun.SqlWrite(s, _strConSql);

                    s = "SELECT * FROM GesMovimenti WHERE mov_nfa = '" + lblMotNfa.Text + "' AND mov_yfa='" + (string)tFat.Rows[0]["fat_yfa"] + "' ORDER BY mov_rfa DESC";
                    tTmp = _clsFun.FillTabSql(TABGESMOV, s, false, _strConSql);
                    if (tTmp.Rows.Count > 0)
                    {
                        foreach (DataRow yy in tTmp.Rows)
                        {
                            if (_clsFun.Numerico(tTmp.Rows[0]["mov_rfa"]))
                            {
                                i = Convert.ToInt16(tTmp.Rows[0]["mov_rfa"]);
                                break;
                            }
                        }
                    }
                }
                else
                {
                    s = _clsFun.SqlInsertRow(TABGESFAT, tFat, x);
                    _clsFun.SqlWrite(s, _strConSql);
                }

                foreach (DataRow y in t.Rows)
                {
                    if (!(Boolean)y["MotExl"] && !(Boolean)y["mot_ann"])
                    {
                        s = "UPDATE GesMovTestate SET ";
                        s += "mot_yfa='" + dtpFatDdo.Value.Year.ToString() + "', ";
                        s += "mot_nfa='" + sNmo + "' ";
                        s += "WHERE ";
                        s += "mot_ymo='" + y["mot_ymo"] + "' AND ";
                        s += "mot_nmo='" + y["mot_nmo"] + "' ";
                        _clsFun.SqlWrite(s, _strConSql);

                        s = "UPDATE GesMovimenti SET ";
                        s += "mov_yfa='" + dtpFatDdo.Value.Year.ToString() + "', ";
                        s += "mov_nfa='" + sNmo + "' ";
                        s += "WHERE ";
                        s += "mov_ymo='" + y["mot_ymo"] + "' AND ";
                        s += "mov_nmo='" + y["mot_nmo"] + "' ";
                        _clsFun.SqlWrite(s, _strConSql);

                        s = "SELECT * FROM GesMovimenti WHERE ";
                        s += "mov_yfa='" + dtpFatDdo.Value.Year.ToString() + "' AND ";
                        s += "mov_nfa='" + sNmo + "' AND ";
                        s += "mov_nmo='" + y["mot_nmo"] + "' ";
                        s += "ORDER BY mov_nmo, mov_rmo";
                        tTmp = _clsFun.FillTabSql(TABGESMOV, s, false, _strConSql);

                        string sMov = "";

                        foreach (DataRow k in tTmp.Rows)
                        {

                            if (sMov != (string)k["mov_nmo"])
                            {
                                sMov = (string)k["mov_nmo"];
                                i++;

                                x = tTmp.NewRow();
                                x["mov_ubi"] = sUbi;
                                x["mov_yfa"] = sYea;
                                x["mov_nfa"] = sNmo;
                                x["mov_rfa"] = i.ToString("0000");
                                x["mov_ymo"] = _clsDef.COD04X;
                                x["mov_nmo"] = _clsDef.COD06X;
                                x["mov_rmo"] = _clsDef.COD04X;
                                x["mov_art"] = "";
                                x["mov_ard"] = "Documento Numero " + (string)y["mot_ndo"] + " del " + ((DateTime)y["mot_ddo"]).ToString("dd/MM/yyyy");
                                x["mov_iva"] = "";
                                x["mov_umi"] = "";
                                x["mov_qta"] = 0;
                                x["mov_cos"] = 0;
                                x["mov_prv"] = 0;
                                x["mov_imp"] = 0;
                                x["mov_ann"] = false;
                                x["mov_day"] = DateTime.Today;
                                x["mov_no1"] = "";

                                s = _clsFun.SqlInsertRow(TABGESMOV, tTmp, x);
                                _clsFun.SqlWrite(s, _strConSql);
                            }

                            i++;

                            s = "UPDATE GesMovimenti SET ";
                            s += "mov_rfa='" + i.ToString("0000") + "' WHERE ";
                            s += "mov_ubi='" + k["mov_ubi"] + "' AND ";
                            s += "mov_ymo='" + k["mov_ymo"] + "' AND ";
                            s += "mov_nmo='" + k["mov_nmo"] + "' AND ";
                            s += "mov_yfa='" + k["mov_yfa"] + "' AND ";
                            s += "mov_nfa='" + k["mov_nfa"] + "' AND ";
                            s += "mov_rmo='" + k["mov_rmo"] + "'";
                            _clsFun.SqlWrite(s, _strConSql);
                        }
                    }
                }

                frmGesDocumento f = new frmGesDocumento();
                f._strMovFat = "F";
                f._strCauTpd = cmbFatTpd.SelectedValue.ToString();
                f._strMftYea = dtpFatDdo.Value.Year.ToString();
                f._strMftNum = sNmo;
                f._strDocSuf = _strSuf;
                f.ShowDialog();
                if (f._strReturn != "")
                {
                    string[] aa = f._strReturn.Split('|');
                    //if (cmbMovFat.SelectedValue.ToString() == "F" && cmbCauTpd.SelectedValue.ToString() == "FV")
                    if (cmbFatTpd.SelectedValue.ToString() == "F")
                    {
                        x["DocStd"] = aa[0];
                        x["DocCfd"] = aa[1];
                        x["DocNdo"] = aa[2];
                    }
                }
            }
        }

        private void cmbFatCli_SelectionChangeCommitted(object sender, EventArgs e)
        {
            if(dgv1.DataSource != null)
                ((DataTable)dgv1.DataSource).Clear();
        }

        private void fatturaDiRiferimentoToolStripMenuItem_Click(object sender, EventArgs e)
        {
            frmSeekDocs f = new frmSeekDocs();
            f._strImpDoc = "S";
            f.ShowDialog();
            if(f._strImpDoc != "")
            {
                string[] a = f._strImpDoc.Split(';');

                if (a.Length > 4 &&  a[0] == "F")
                {
                    if (MessageBox.Show("Confermi l'aggregazione dei documenti al movimento fattura numero " + a[3] + "?", "RAGGRUPPAMENTO SU FATTU GIA' PRESENTE", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                    {
                        lblMotNfa.Text = a[3];
                        lblMotNfa.BackColor = Color.Red;
                    }

                    //string sNfa = a[3];

                    //if (MessageBox.Show("Confermi l'aggregazione dei documenti al movimento fattura numero " + sNfa + "?", "RAGGRUPPAMENTO SU FATTU GIA' PRESENTE", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                    //{
                    //    CurrencyManager cm = dgv1.BindingContext[dgv1.DataSource, dgv1.DataMember] as CurrencyManager;
                    //    if (cm.Position >= 0)
                    //    {
                    //        DataRowView r = cm.List[dgv1.CurrentRow.Index] as DataRowView;
                    //        DataRow x = r.Row;
                    //        x["mot_nfa"] = sNfa;
                    //    }

                    //    DataTable t = (DataTable)dgv1.DataSource;

                    //    foreach (DataRow y in t.Rows)
                    //    {
                    //        if (!(Boolean)y["mot_ann"] && !(Boolean)y["MotExl"])
                    //            y["mot_nfa"] = sNfa;
                    //    }
                    //}
                }

            }
        }


    }
}
