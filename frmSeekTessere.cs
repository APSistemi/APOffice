using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;

namespace APOffice
{
    public partial class frmSeekTessere : Form
    {
        private const string TABANATES = "AnaTessere";
        private const string TABTABCAU = "TabCauCassa";
        private const string TABTABSTD = "TabStatoDivulgazioni";

        clsDefine _clsDef = new clsDefine();
        clsFuncs _clsFun = new clsFuncs();

        public string _strCod = "";
        public string _strDes = "";

        public bool _bolNew = true;
        public string _strSqlWhe = "";
        public bool _bolAnagra = true;
        public bool _bolSelect = false;

        private string _strConSql = "";

        public DataTable _tabTmp = new DataTable();

        private DateTime _dayDti = new DateTime();
        private DateTime _dayDtf = new DateTime();

        public frmSeekTessere()
        {
            InitializeComponent();
            new clsGesGraph().SetGraph(this, 0);
        }

        private void frmSeekFidelity_Load(object sender, EventArgs e)
        {
            ApplyModernUi();
            _strConSql = _clsFun.ConSql("");

            SetDgv1();

            _dayDti = _clsDef.DAYOUT;
            _dayDtf = _clsDef.DAYOUT;

            DataTable t = new clsQuery().Campagna();
            if (t.Rows.Count > 0)
            {
                _dayDti = (DateTime)t.Rows[0]["tab_dti"];
                _dayDtf = (DateTime)t.Rows[0]["tab_dtf"];
            }
            else
            {
                MessageBox.Show("Campagna non definita in tabella!");
            }

            txtDes.Select();

            clsUiIcons.RestoreFormBounds(this);
            clsUiIcons.RestoreGridColumnWidths(dgv1, "frmSeekTessere_dgv1");
        }

        private void frmSeekTessere_FormClosing(object sender, FormClosingEventArgs e)
        {
            clsUiIcons.SaveFormBounds(this);
            clsUiIcons.SaveGridColumnWidths(dgv1, "frmSeekTessere_dgv1");
        }

        private void ApplyModernUi()
        {
            try
            {
                if (menuStrip1 != null)
                {
                    menuStrip1.Renderer = clsUiIcons.GetModernMenuRenderer();
                    if (esciToolStripMenuItem != null)
                        esciToolStripMenuItem.Image = clsUiIcons.GetIcon("exit", 16);
                    if (excelToolStripMenuItem != null)
                        excelToolStripMenuItem.Image = clsUiIcons.GetIcon("excel", 16);
                }

                this.BackColor = Color.FromArgb(243, 244, 246);

                if (btnAll != null)
                {
                    clsUiIcons.StyleStatButton(btnAll, "Tutti", "list", 18,
                        Color.FromArgb(239, 246, 255), Color.FromArgb(191, 219, 254),
                        Color.FromArgb(96, 165, 250), Color.FromArgb(30, 58, 138), Color.FromArgb(29, 78, 216));
                }

                if (btnNew != null)
                {
                    clsUiIcons.StyleStatButton(btnNew, "Nuovo", "plus", 18,
                        Color.FromArgb(236, 253, 245), Color.FromArgb(167, 243, 208),
                        Color.FromArgb(52, 211, 153), Color.FromArgb(6, 78, 59), Color.FromArgb(5, 150, 105));
                }

                if (btnPrn != null)
                {
                    clsUiIcons.StyleStatButton(btnPrn, "Stampa", "print", 16,
                        Color.FromArgb(254, 252, 232), Color.FromArgb(254, 240, 138),
                        Color.FromArgb(250, 204, 21), Color.FromArgb(113, 63, 18), Color.FromArgb(202, 138, 4));
                }

                if (btnPos != null)
                {
                    clsUiIcons.StyleStatButton(btnPos, "Invio alle casse", "send", 16,
                        Color.FromArgb(240, 253, 250), Color.FromArgb(153, 246, 228),
                        Color.FromArgb(45, 212, 191), Color.FromArgb(19, 78, 74), Color.FromArgb(13, 148, 136));
                }

                if (dgv1 != null)
                {
                    clsUiIcons.StyleDataGridView(dgv1);
                }
            }
            catch (Exception ex)
            {
                _clsFun.ErrorLog("frmSeekTessere.ApplyModernUi", ex.Message);
            }
        }

        private void frmSeekTessere_KeyDown(object sender, KeyEventArgs e)
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
            DataTable tSta = _clsFun.FillTabSql(TABTABSTD, "SELECT * FROM " + TABTABSTD + " WHERE tab_var=1", false, _strConSql);

            dgv1.AutoGenerateColumns = false;
            dgv1.AllowUserToAddRows = false;
            dgv1.ReadOnly = false;
            dgv1.AllowUserToDeleteRows = false;

            DataGridViewTextBoxColumn cTbc;
            DataGridViewCheckBoxColumn cCbc;
            DataGridViewComboBoxColumn cCmb;

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "tes_cod";
            cTbc.Name = "Codice";
            cTbc.Width = 110;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = true;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "tes_des";
            cTbc.Name = "Descrizione";
            cTbc.Width = 260;
            cTbc.AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = true;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "TesImp";
            cTbc.Name = "Acquisti";
            cTbc.Width = 80;
            cTbc.ValueType = typeof(decimal);
            cTbc.ReadOnly = true;
            cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            cTbc.DefaultCellStyle.Format = "#,##0.00";
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "TesPun";
            cTbc.Name = "Punti";
            cTbc.Width = 70;
            cTbc.ValueType = typeof(int);
            cTbc.ReadOnly = true;
            cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            cTbc.DefaultCellStyle.Format = "#,##0";
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "TesGrd";
            cTbc.Name = "Gruppo";
            cTbc.Width = 90;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = true;
            dgv1.Columns.Add(cTbc);

            cCbc = new DataGridViewCheckBoxColumn();
            cCbc.ValueType = typeof(Boolean);
            cCbc.DataPropertyName = "tes_ann";
            cCbc.Name = "Ann.";
            cCbc.Width = 45;
            cCbc.ReadOnly = true;
            dgv1.Columns.Add(cCbc);

            cCmb = new DataGridViewComboBoxColumn();
            cCmb.DataPropertyName = "TesInv";
            cCmb.Name = "Stato";
            cCmb.Width = 110;
            cCmb.DataSource = tSta;
            cCmb.ValueMember = "tab_cod";
            cCmb.DisplayMember = "tab_des";
            cCmb.ReadOnly = true;
            cCmb.DisplayStyle = DataGridViewComboBoxDisplayStyle.Nothing;
            cCmb.DefaultCellStyle.Font = new Font("Segoe UI", 9F, GraphicsUnit.Point);
            dgv1.Columns.Add((DataGridViewColumn)cCmb);
        }

        private void btnAll_Click(object sender, EventArgs e)
        {
            TabSeek("ALL");
        }

        private void txtCod_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Return)
                TabSeek("COD");
        }

        private void txtDes_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Return)
                TabSeek("DES");
        }

        private void TabSeek(string strTip)
        {
            DataRow[] j;
            string p = TABANATES;
            string s = "";
            string w = "";

            if (_strSqlWhe != "" && strTip == "")
            {
                w = _strSqlWhe;
            }
            else if (strTip == "ALL")
            {
                w = strTip;
            }
            else if (strTip == "DES")
            {
                if (txtDes.Text != "")
                    w = "tes_des LIKE '%" + txtDes.Text.Trim() + "%'";
            }
            else if (strTip == "COD")
            {
                if (txtCod.Text != "")
                {
                    txtCod.Text = txtCod.Text.PadLeft(13, Convert.ToChar("0"));
                    w = "tes_cod='" + txtCod.Text + "'";
                }
            }
            if (w != "")
            {
                if (!chkAnn.Checked)
                    w += " AND tes_ann=0 ";

                _strSqlWhe = s;
                if (strTip == "ALL")
                {
                    s = "SELECT * FROM " + p + " ";
                    if (!chkAnn.Checked)
                        s += "WHERE tes_ann=0 ";
                    s += "ORDER BY tes_des";
                }
                else
                {
                    s = "SELECT * FROM " + p + " ";
                    if(w != "")
                    s += "WHERE " + w + " ";
                    s += " ORDER BY tes_des";
                }
                DataTable t = _clsFun.FillTabSql(TABANATES, s, false, _strConSql);

                if(t.Rows.Count > 0)
                {
                    t.Columns.Add(new DataColumn()
                    {
                        DataType = Type.GetType("System.Decimal"),
                        ColumnName = "TesPun",
                        Caption = "Punti",
                        ReadOnly = false,
                        DefaultValue = (Decimal)0
                    });

                    t.Columns.Add(new DataColumn()
                    {
                        DataType = Type.GetType("System.Decimal"),
                        ColumnName = "TesImp",
                        Caption = "Importo",
                        ReadOnly = false,
                        DefaultValue = (Decimal)0
                    });

                    if (_dayDti <= DateTime.Today && _dayDtf >= DateTime.Today)
                    {
                        //s = "SELECT * FROM " + TABTABCAU;
                        //DataTable tCau = _clsFun.FillTabSql(TABTABCAU, s, false, _strConSql);

                        //DateTime dDti = DateTime.Today;
                        //DateTime dDtf = DateTime.Today;

                        //p = "TabFidCampagne";
                        //s = "SELECT * FROM " + p + " WHERE tab_ann=0 ORDER BY tab_dti DESC";
                        //t = _clsFun.FillTabSql(p, s, false, _strConSql);
                        //if(t.Rows.Count > 0)
                        //{
                        //    dDti = (DateTime)t.Rows[0]["tab_dti"];
                        //    dDtf = (DateTime)t.Rows[0]["tab_dti"];
                        //}

                        s = "SELECT * FROM TabFidGruppi";
                        DataTable tGru = _clsFun.FillTabSql("TabGru", s, false, _strConSql);

                        if(t.Columns.IndexOf("TesGrd") < 0)
                        {
                            t.Columns.Add(new DataColumn()
                            {
                                DataType = Type.GetType("System.String"),
                                ColumnName = "TesGrd",
                                Caption = "Descrizione",
                                MaxLength = 50,
                                ReadOnly = false,
                                DefaultValue = (String)""
                            });
                        }

                        s = "SELECT ";
                        s += "vet_fid, ";
                        s += "SUM(CASE WHEN vep_tip = 'PAG' THEN vep_imp ELSE 0 END) AS vep_imp, ";
                        s += "SUM(CASE WHEN vep_tip = 'PUM' THEN ABS(vep_imp) * -1 ELSE vep_imp END) AS vep_pun, ";
                        s += "SUM(vet_imp) AS vet_imp ";
                        s += "FROM ";
                        s += "(SELECT GesNegVet.vet_fid, GesNegVep.vep_tip, GesNegVep.vep_imp, GesNegVet.vet_imp ";
                        s += "FROM GesNegVet ";
                        s += "LEFT OUTER JOIN GesNegVep ON ";
                        s += "GesNegVet.vet_sco = GesNegVep.vep_sco AND ";
                        s += "GesNegVet.vet_pos = GesNegVep.vep_pos AND ";
                        s += "GesNegVet.vet_ora = GesNegVep.vep_ora AND ";
                        s += "GesNegVet.vet_day = GesNegVep.vep_day AND ";
                        s += "GesNegVet.vet_neg = GesNegVep.vep_neg AND ";
                        s += "GesNegVet.vet_cau = GesNegVep.vep_cau ";
                        s += "WHERE ";

                        s += "(GesNegVet.vet_day >= " + _clsFun.DaySql(_dayDti) + " AND ";

                        s += "GesNegVep.vep_cau = 'PUN' AND GesNegVep.vep_ean <> '') OR ";

                        s += "(GesNegVet.vet_day >= " + _clsFun.DaySql(_dayDti) + " AND ";

                        s += "LEFT(GesNegVep.vep_tip, 2) = 'PU' AND GesNegVep.vep_ean = '' OR GesNegVep.vep_ean = '" + _clsDef.PUNTITESSERAX + "')) AS derivedtbl_1 ";
                        //s += "LEFT(GesNegVep.vep_tip, 2) = 'PU')) AS derivedtbl_1 ";

                        s += "GROUP BY vet_fid ";

                        DataTable tMon = _clsFun.FillTabSql("TabMot", s, false, _clsFun.ConSql("3"));
                        DataColumn[] Key = new DataColumn[1] 
                        { 
                            tMon.Columns["vet_fid"] 
                        };
                        tMon.PrimaryKey = Key;

                        foreach (DataRow y in t.Rows)
                        {
                            j = tMon.Select("vet_fid='" + (string)y["tes_cod"] + "'");
                            if (j.Length > 0)
                            {
                                y["TesImp"] = (decimal)j[0]["vet_imp"];
                                y["TesPun"] = (decimal)j[0]["vep_pun"];
                            }

                            j = tGru.Select("tab_cod='" + y["tes_gru"] + "'");
                            if(j.Length > 0)
                                y["TesGrd"] = (string)j[0]["tab_des"];
                        }
                    }
                }

                t.Columns.Add(new DataColumn()
                {
                    DataType = Type.GetType("System.String"),
                    ColumnName = "TesInv",
                    Caption = "Invio",
                    MaxLength = 1,
                    ReadOnly = false,
                    DefaultValue = (String)"5"
                });

                t.DefaultView.AllowDelete = false;
                t.DefaultView.AllowEdit = false;
                t.DefaultView.AllowNew = false;
                dgv1.DataSource = t;
                dgv1.Focus();
            }
            else if (strTip != "")
            {
                MessageBox.Show("Tessera non trovata");
            }
        }

        private void btnNew_Click(object sender, EventArgs e)
        {
            frmAnaTessera f = new frmAnaTessera();
            f._strCod = _clsDef.CODNEW;
            f.ShowDialog();

            if (f._strCod != "" && f._strCod != _clsDef.CODNEW)
            {
                txtCod.Text = f._strCod;
                TabSeek("COD");
            }
        }

        private void dgv1_DoubleClick(object sender, EventArgs e)
        {
            Scelto();
        }

        private void dgv1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.ColumnIndex == 6)
            {
                CurrencyManager cm = dgv1.BindingContext[dgv1.DataSource, dgv1.DataMember] as CurrencyManager;
                if (cm.Position >= 0)
                {
                    DataRowView r = cm.List[dgv1.CurrentRow.Index] as DataRowView;
                    DataRow x = r.Row;

                    string s = (string)x["TesInv"];
                    if (s == "5")
                        x["TesInv"] = "0";
                    else if (s == "0")
                        x["TesInv"] = "5";

                }
            }
        }

        private void dgv1_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Return)
                Scelto();
        }

        private void Scelto()
        {
            CurrencyManager cm = dgv1.BindingContext[dgv1.DataSource, this.dgv1.DataMember] as CurrencyManager;
            if (cm.Position >= 0)
            {
                DataRowView r = cm.List[dgv1.CurrentRow.Index] as DataRowView;
                DataRow x = r.Row;

                _strCod = Convert.ToString(x["tes_cod"]);
                _strDes = Convert.ToString(x["tes_des"]);

                _tabTmp = new clsGenTabTmp().TabTmpTes("AnaTes");
                DataRow k = _tabTmp.NewRow();
                k["tmp_tes"] = _strCod;
                k["tmp_ted"] = _strDes;
                _tabTmp.Rows.Add(k);

                if (_bolAnagra && !_bolSelect)
                {
                    frmAnaTessera f = new frmAnaTessera();
                    f._tabTmp = _tabTmp.Copy();
                    f.ShowDialog();
                    if (f._bolCodNew && f._strCod != _clsDef.CODNEW)
                        TabSeek("DES");
                    else if (f._bolCodSostituito)
                    {
                        txtCod.Text = f._strCod;
                        TabSeek("COD");
                    }
                    if (f._strDes != "")
                    {
                        x["tes_des"] = f._strDes;
                        x["tes_gru"] = f._strGru;
                    }
                }
                else
                {
                    if ((Boolean)x["tes_ann"])
                        MessageBox.Show("Tessera annullata");
                    else
                        Esci();
                }
            }
            else
                Esci();
        }

        private void btnPos_Click(object sender, EventArgs e)
        {
            if (dgv1.DataSource == null || ((DataTable)dgv1.DataSource).Rows.Count == 0)
            {
                MessageBox.Show("Righe non presenti!");
            }
            else
            {
                if (MessageBox.Show("Se continui verranno aggiornati i saldi in cassa delle tessere selezionate, confermi?", this.Text, MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                    Tes2Pos();
            }
        }

        private void Tes2Pos()
        {
            string sMsg = "";
            DataRow[] j;

            DataTable t = (DataTable)dgv1.DataSource;

            DataTable tTmp = t.Clone();

            foreach (DataRow y in t.Rows)
            {
                if (((string)y["tes_gru"]).Trim() == "")
                    sMsg += "Tessera " + (string)y["tes_cod"] + " con GRUPPO non definito " + _clsDef.CRLF;
                else
                    tTmp.ImportRow(y);
            }

            if (tTmp.Rows.Count > 0)
            {
                frmGesVarPos f = new frmGesVarPos();
                f._tabFid = tTmp;
                f.ShowDialog();

                foreach (DataRow y in t.Rows)
                {
                    j = tTmp.Select("tes_cod='" + y["tes_cod"] + "'");
                    if (j.Length > 0)
                        y["TesInv"] = "5";
                }
            }

            if(sMsg != "")
                MessageBox.Show(sMsg, "CONTROLLO TESSERE", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        }

        private void chkAll_Click(object sender, EventArgs e)
        {
            string sSta = "5";
            if (chkAll.Checked)
                sSta = "0";
            DataTable t = (DataTable)dgv1.DataSource;
            foreach(DataRow y in t.Rows)
                y["TesInv"] = sSta;
        }

        private void excelToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (dgv1.DataSource == null)
                MessageBox.Show("Non ci sono tessere estratte!", "ESTRAZIONE TESSERE", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            else
            {
                //Excel1();
                Csv1();
            }
        }

        private void Excel1()
        {
            DataRow[] j;
            DataTable t = new DataTable();
            DataView v = new DataView((DataTable)dgv1.DataSource, "", "tes_cod", DataViewRowState.CurrentRows);
            t = v.Table.Clone();

            string s = "";
            decimal d = 0;

            foreach (DataRowView r in v)
            {
                t.ImportRow(r.Row);
            }

            string sTit = DateTime.Today.ToString("dd/MM/yyyy") + " - Estrazione tessere";
            string sFil = "";

            s = "";
            s += "tes_cod, Codice, 100, StringLiteral;";
            s += "tes_des, Descrizione, 160, StringLiteral;";
            s += "TesPun, Punti, 50, Integer;";
            s += "TesImp, Acquisti, 50, Decimal2;";
            s += "TesGrd, Gruppo, 80, StringLiteral;";

            for (int i = 0; i <= 100; i++)
            {
                sFil = "C:\\APproject\\XLS\\" + DateTime.Now.ToString("yyyyMMdd") + "_Tessere.xls";
                if (File.Exists(sFil))
                {
                    try
                    {
                        File.Delete(sFil);
                        break;
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message);
                    }

                }
                else
                    break;
            }

            DataSet ds = new DataSet();

            DataTable tTit = (new clsExcel()).FillTitStr(s, "TabTit");

            ds.Tables.Add(tTit);
            ds.Tables.Add(t);

            (new clsExcel()).exportToExcel(ds, sFil, sTit, "", true);

            ds.Tables.Remove(t);
            ds.Clear();
            ds.Dispose();
        }

        private void Csv1()
        {
            DataRow[] j;
            DataTable t = new DataTable();
            DataView v = new DataView((DataTable)dgv1.DataSource, "", "tes_cod", DataViewRowState.CurrentRows);
            t = v.Table.Clone();

            string s = "";
            decimal d = 0;

            foreach (DataRowView r in v)
            {
                t.ImportRow(r.Row);
            }

            string sTit = DateTime.Today.ToString("dd/MM/yyyy") + " - Estrazione tessere";
            string sFil = "";

            s = "";
            s += "tes_cod, Codice, 100, StringLiteral;";
            s += "tes_des, Descrizione, 160, StringLiteral;";
            s += "TesPun, Punti, 50, Integer;";
            s += "TesImp, Acquisti, 50, Decimal2;";
            s += "TesGrd, Gruppo, 80, StringLiteral;";

            string sFld = s;

            (new clsExcel()).exportToCsv1(sFld, t, sFil, sTit, "");

        }

    }
}
