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
    public partial class frmSeekOrdFor : Form
    {
        clsDefine _clsDef = new clsDefine();
        clsFuncs _clsFun = new clsFuncs();
        clsQuery _clsQry = new clsQuery();

        private const string TABGESORT = "GesOrdForTestate";
        private const string TABGESORF = "GesOrdForRighe";

        private string _strConSql = "";

        public frmSeekOrdFor()
        {
            InitializeComponent();
            new clsGesGraph().SetGraph(this, 0);
        }

        private void frmSeekOrd_Load(object sender, EventArgs e)
        {
            ApplyModernUi();
            _strConSql = _clsFun.ConSql("");
            txtOftYea.Text = DateTime.Today.Year.ToString();
            SetDgv1();
            FillTabs();
            FillDati();

            clsUiIcons.RestoreFormBounds(this);
            clsUiIcons.RestoreGridColumnWidths(dgv1, "frmSeekOrdFor_dgv1");
        }

        private void frmSeekOrdFor_FormClosing(object sender, FormClosingEventArgs e)
        {
            clsUiIcons.SaveFormBounds(this);
            clsUiIcons.SaveGridColumnWidths(dgv1, "frmSeekOrdFor_dgv1");
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
                    if (importApPhonToolStripMenuItem != null)
                        importApPhonToolStripMenuItem.Image = clsUiIcons.GetIcon("smartphone", 16);
                    if (importSuddivisoToolStripMenuItem != null)
                        importSuddivisoToolStripMenuItem.Image = clsUiIcons.GetIcon("file_text", 16);
                    if (cancellaToolStripMenuItem != null)
                        cancellaToolStripMenuItem.Image = clsUiIcons.GetIcon("trash", 16);
                }

                this.BackColor = Color.FromArgb(243, 244, 246);

                if (btnNew != null)
                {
                    clsUiIcons.StyleStatButton(btnNew, "Nuovo", "plus", 16,
                        Color.FromArgb(236, 253, 245), Color.FromArgb(167, 243, 208),
                        Color.FromArgb(52, 211, 153), Color.FromArgb(6, 78, 59), Color.FromArgb(5, 150, 105));
                }

                if (dgv1 != null)
                {
                    clsUiIcons.StyleDataGridView(dgv1);
                }
            }
            catch (Exception ex)
            {
                _clsFun.ErrorLog("frmSeekOrdFor.ApplyModernUi", ex.Message);
            }
        }

        private void frmSeekOrdFor_KeyDown(object sender, KeyEventArgs e)
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

        private void FillTabs()
        {
            DataRow x;
            string p = "";
            string s = "";
            DataTable t = new DataTable();

            p = "TabForImport";
            s = "SELECT * FROM TabForImport WHERE (tab_nor IS NULL OR tab_nor = 0)";
            t = _clsFun.FillTabSql(p, s, false, _strConSql);

            x = t.NewRow();
            x["tab_cod"] = "";
            x["tab_des"] = "  Non definito";
            t.Rows.InsertAt(x, 0);

            cmbOftFor.DataSource = t;
            cmbOftFor.DisplayMember = "tab_des";
            cmbOftFor.ValueMember = "tab_cod";
            if(t.Rows.Count > 0)
                cmbOftFor.SelectedIndex = 0;
            else
                MessageBox.Show("Non ci sono fornitori attivi", "CONTROLLO FORNITORI", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }

        private void FillDati()
        {
            string s = "";
            s = "SELECT ";
            s += "GesOrdForTestate.*, ";
            s += "AnaFornitori.for_des AS ForDes, ";
            s += "TabStatoDivulgazioni.tab_des AS OftSta ";
            s += "FROM GesOrdForTestate ";
            s += "LEFT OUTER JOIN AnaFornitori ON GesOrdForTestate.oft_for = AnaFornitori.for_cod ";
            s += "LEFT OUTER JOIN TabStatoDivulgazioni ON GesOrdForTestate.oft_sta = TabStatoDivulgazioni.tab_cod ";
            s += "WHERE ";
            s += "oft_yea='" + txtOftYea.Text + "' ";
            s += "ORDER BY oft_num DESC";
            DataTable t = _clsFun.FillTabSql("OFT", s, false, _strConSql);

            dgv1.DataSource = t;
        }

        private void SetDgv1()
        {
            dgv1.AutoGenerateColumns = false;
            dgv1.AllowUserToAddRows = false;
            dgv1.ReadOnly = false;
            dgv1.AllowUserToDeleteRows = false;

            DataGridViewTextBoxColumn cTbc;

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "oft_day";
            cTbc.Name = "Data";
            cTbc.Width = 90;
            cTbc.ValueType = typeof(string);
            cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            cTbc.ReadOnly = true;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "ForDes";
            cTbc.Name = "Fornitore";
            cTbc.Width = 260;
            cTbc.AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = true;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "oft_num";
            cTbc.Name = "Numero";
            cTbc.Width = 80;
            cTbc.ValueType = typeof(string);
            cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            cTbc.ReadOnly = true;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "OftSta";
            cTbc.Name = "Stato";
            cTbc.Width = 140;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = true;
            dgv1.Columns.Add(cTbc);
        }

        private void txtOftYea_KeyDown(object sender, KeyEventArgs e)
        {
            if(e.KeyCode == Keys.Return)
                FillDati();
        }

        private void btnNew_Click(object sender, EventArgs e)
        {
            Boolean b = false;
            if (cmbOftFor.SelectedValue.ToString() == "")
                MessageBox.Show("Fornitore non selezionato!", "CONTROLLO FORNITORE", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            else
                b = true;
            if (b)
            {
                frmGesOrdFornitori f = new frmGesOrdFornitori();
                f._strOftYea = txtOftYea.Text;
                f._strOftFor = cmbOftFor.SelectedValue.ToString().Trim();
                f._strOftNum = _clsDef.CODNEW;
                f.ShowDialog();
                FillDati();
            }
        }

        private void dgv1_DoubleClick(object sender, EventArgs e)
        {
            Scelto();
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

                frmGesOrdFornitori f = new frmGesOrdFornitori();
                f._strOftYea = (string)x["oft_yea"];
                f._strOftFor = ((string)x["oft_for"]).Trim();
                f._strOftNum = (string)x["oft_num"];
                f.ShowDialog();
                if (f._strOftStd != "")
                    x["OftSta"] = f._strOftStd.Trim();
                //FillDati();
            }
        }

        private void importSuddivisoToolStripMenuItem_Click(object sender, EventArgs e)
        {
            ImpTerm();
        }

        private void ImpTerm()
        {
            DataRow[] j;
            string s = "";
            Boolean b = true;
            DataTable tMsg = new clsGenTabTmp().TabTmpVenErrori("TmpMsg");

            s = "SELECT * FROM GesOrdForTestate WHERE oft_num='" + _clsDef.COD06X + "'";
            DataTable tOrt = _clsFun.FillTabSql(TABGESORT, s, false, _strConSql);

            s = _clsFun.ParGet(clsDefine.enuParametri.ParOrdForDividi, _strConSql);
            if (!_clsFun.Numerico(s))
                b = false; 
            string[] aFor = s.Split(',');
            if (aFor.Length == 0)
                b = false;
            if(!b)
                MessageBox.Show("Funzione non attiva!");
            else
            {
                for (int i1 = 0; i1 < aFor.Length; i1++)
                    aFor[i1] = Convert.ToString(aFor[i1]).PadLeft(5, Convert.ToChar('0'));

                s = "SELECT * ";
                s += "FROM GesOrdForRighe ";
                s += "WHERE ";
                s += "orf_yea='" + txtOftYea.Text + "' AND orf_num='" + _clsDef.COD06X + "'";
                DataTable tOrf = _clsFun.FillTabSql(TABGESORF, s, false, _strConSql);
                DataTable tOrd = tOrf.Clone();

                tOrd.Columns.Add(new DataColumn()
                {
                    DataType = Type.GetType("System.String"),
                    ColumnName = "tmp_for",
                    Caption = "Fornitore",
                    MaxLength = 6,
                    ReadOnly = false,
                    DefaultValue = (String)""
                });

                DataTable t;
                DataTable tLia;
                DataRow x;
                frmGesImpTerm f = new frmGesImpTerm();
                f.ShowDialog();
                if (f._tabImp != null)
                {
                    progressBar1.Visible = true;
                    progressBar1.Value = 0;
                    progressBar1.Maximum = f._tabImp.Rows.Count;
                    progressBar1.Minimum = 0;

                    foreach (DataRow y in f._tabImp.Rows)
                    {
                        progressBar1.Increment(1);
                        System.Windows.Forms.Application.DoEvents();

                        tLia = _clsQry.SeekArtEanCosMin(aFor, y, tMsg);

                        if (tLia.Rows.Count > 0)
                        {
                            x = tOrd.NewRow();
                            x["tmp_for"] = tLia.Rows[0]["lia_for"];
                            x["orf_nri"] = "";
                            x["orf_yea"] = txtOftYea.Text;
                            x["orf_num"] = "";
                            x["orf_day"] = DateTime.Today;
                            x["orf_arf"] = tLia.Rows[0]["lia_arf"];
                            x["orf_ean"] = (string)y["ord_ean"];
                            x["orf_art"] = tLia.Rows[0]["lia_art"];
                            x["orf_qta"] = (decimal)y["ord_qta"];
                            x["orf_cos"] = tLia.Rows[0]["lia_cos"];
                            x["orf_cxp"] = tLia.Rows[0]["lia_cxp"];
                            x["orf_pxc"] = tLia.Rows[0]["lia_pxc"];
                            x["orf_prv"] = 0;
                            x["orf_umi"] = tLia.Rows[0]["art_umi"];
                            x["orf_ard"] = tLia.Rows[0]["art_des"];
                            tOrd.Rows.Add(x);
                        }
                    }

                    progressBar1.Visible = false;

                }

                frmGesOrdForMisto ff = new frmGesOrdForMisto();
                ff._tabOrd = tOrd.Copy();
                ff.ShowDialog();
                tOrd = ff._tabOrd;

                if(tOrd.Rows.Count > 0)
                {
                    DataView v = new DataView(tOrd, "", "tmp_for", DataViewRowState.CurrentRows);

                    string sFor = "";
                    string sNum = "";
                    int iRig = 0;

                    foreach(DataRowView r in v)
                    {
                        if((string)r["tmp_for"] != sFor)
                        {
                            sNum = _clsFun.NewNum(txtOftYea.Text, clsDefine.enuNumeratori.NumGesOrdFor, 5, _strConSql);
                            sFor = (string)r["tmp_for"];
                            x = tOrt.NewRow();
                            x["oft_yea"] = txtOftYea.Text;
                            x["oft_num"] = sNum;
                            x["oft_for"] = sFor;
                            x["oft_day"] = DateTime.Today;
                            x["oft_sta"] = _clsDef.DIVDAD;
                            tOrt.Rows.Add(x);
                            s = _clsFun.SqlInsertRow(TABGESORT, tOrt, x);
                            _clsFun.SqlWrite(s, _strConSql);
                            iRig = 0;
                        }

                        iRig++;

                        x = tOrf.NewRow();

                        foreach (DataColumn c in tOrf.Columns)
                            x[c.ColumnName] = r[c.ColumnName];

                        x["orf_nri"] = iRig.ToString("000");
                        x["orf_yea"] = txtOftYea.Text;
                        x["orf_num"] = sNum;
                        x["orf_day"] = DateTime.Today;

                        tOrf.Rows.Add(x);
                        s = _clsFun.SqlInsertRow(TABGESORF, tOrf, x);
                        _clsFun.SqlWrite(s, _strConSql);
                    }
                }
                if (tMsg.Rows.Count > 0)
                    new clsGenPdfVenErr().PrnPdfOrfErr(tMsg);
                FillDati();
            }
        }

        private void importApPhonToolStripMenuItem_Click(object sender, EventArgs e)
        {
            string s = "SELECT * FROM TabForImport WHERE tab_tip='FTPDIV'";
            DataTable t = _clsFun.FillTabSql("TabForImp", s, true, _strConSql);
            if(t.Rows.Count > 0)
                new frmGesForDivulgazioni().ForDownLoad();
            else
                MessageBox.Show("Opzione non attiva!", "IMPORTAZIONE DA APSHOP", MessageBoxButtons.OK, MessageBoxIcon.Question);
            //string s = "SELECT * FROM TabForImport WHERE tab_tip='FTPDIV'";
            //DataTable t = _clsFun.FillTabSql("TabForImport", s, true, _strConSql);
            //if (t.Rows.Count > 0)
            //{
            //    ArrayList a = new ArrayList();
            //    a.Add((string)t.Rows[0]["tab_fil"]);
            //    a.Add((string)t.Rows[0]["tab_ord"]);
            //    s = (string)t.Rows[0]["tab_ftp"];
            //    s = _clsQry.ApPhoneDiv2TmpDiv(s, a);
            //}
        }

        private void cancellaToolStripMenuItem_Click(object sender, EventArgs e)
        {
            CancellaOrdine();
        }

        private void CancellaOrdine()
        {
            CurrencyManager cm = dgv1.BindingContext[dgv1.DataSource, this.dgv1.DataMember] as CurrencyManager;
            if (cm.Position >= 0)
            {
                DataRowView r = cm.List[dgv1.CurrentRow.Index] as DataRowView;
                DataRow x = r.Row;

                string s = "";

                string sYea = (string)x["oft_yea"];
                string sFor = ((string)x["oft_for"]).Trim();
                string sNum = (string)x["oft_num"];

                s = "SELECT * FROM GesOrdForTestate WHERE oft_yea='" + sYea + "' AND oft_num='" + sNum + "'";
                DataTable t = _clsFun.FillTabSql("GesOrdForTestate", s, true, _strConSql);
                if(t.Rows.Count > 0)
                {
                    if((string)t.Rows[0]["oft_sta"] != "4")
                        MessageBox.Show("Lo stato dell'ordine deve essere CANCELLATO!", "CANCELLAZIONE ORDINE", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    else
                    {
                        if (MessageBox.Show("Cancellazione dell'ordine n. " + sNum + ", confermi?", "CANCELLAZIONE ORDINE", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                        {
                            s = "DELETE FROM GesOrdForRighe WHERE orf_yea='" + sYea + "' AND orf_num='" + sNum + "'";
                            _clsFun.SqlWrite(s, _strConSql);

                            s = "DELETE FROM GesOrdForTestate WHERE oft_yea='" + sYea + "' AND oft_num='" + sNum + "'";
                            _clsFun.SqlWrite(s, _strConSql);

                            FillDati();
                        }
                   }
                }
            }

        }
 
    }
}
