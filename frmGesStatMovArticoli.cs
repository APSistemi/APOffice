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
    public partial class frmGesStatMovArticoli : Form
    {
        private const string TABSTAVET = "GesNegVet";
        private const string TABSTAVEP = "GesNegVep";
        private const string TABSTAVEN = "GesNegVen";

        clsDefine _clsDef = new clsDefine();
        clsFuncs _clsFun = new clsFuncs();
        clsQuery _clsQry = new clsQuery();
        clsVariazioni _clsVar = new clsVariazioni();

        private string _strConSql = "";
        private string _strConSqlSta = "";

        public frmGesStatMovArticoli()
        {
            InitializeComponent();
            new clsGesGraph().SetGraph(this, 0);
            _strConSql = _clsFun.ConSql("");
            _strConSqlSta = _clsFun.ConSql("3");
        }

        private void frmGesStaMovArticoli_Load(object sender, EventArgs e)
        {
            dtpIni.Value = new DateTime(DateTime.Today.Year, 1, 1);
            SetDgv1();
        }

        private void esciToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Esci();
        }

        private void frmGesStaMovArticoli_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Escape)
                Esci();
        }
        private void Esci()
        {
            this.Close();
        }

        private void btnEstrai_Click(object sender, EventArgs e)
        {
            FillDati();
        }

        private void SetDgv1()
        {
            //DataTable tSta = _clsFun.FillTabSql("TabStato", "SELECT * FROM TabStato", false, _strConSql);

            dgv1.AutoGenerateColumns = false;
            //dgv1.VirtualMode = true;
            //dgv1.Dock = DockStyle.Fill;
            dgv1.AllowUserToAddRows = false;
            dgv1.ReadOnly = false;
            dgv1.AllowUserToDeleteRows = false;
            //dgv1.DisplayedRowCount() = true;
            dgv1.SelectionMode = DataGridViewSelectionMode.CellSelect;

            DataGridViewTextBoxColumn cTbc;
            //DataGridViewCheckBoxColumn cCbc;
            DataGridViewComboBoxColumn cCmb;

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "Lv3Des";
            cTbc.Name = "ECR";
            cTbc.Width = 120;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = true;
            //cTbc.Selected = true;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "RepDes";
            cTbc.Name = "Reparto";
            cTbc.Width = 120;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = true;
            //cTbc.Selected = true;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "art_cod";
            cTbc.Name = "Articolo";
            cTbc.Width = 70;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = true;
            //cTbc.Selected = true;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "art_des";
            cTbc.Name = "Descrizione";
            cTbc.Width = 190;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = true;
            //cTbc.Selected = true;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "art_umi";
            cTbc.Name = "UM";
            cTbc.Width = 30;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = true;
            //cTbc.Selected = true;
            dgv1.Columns.Add(cTbc);

            //cCmb = new DataGridViewComboBoxColumn();
            //cCmb.DataPropertyName = "art_sta";
            //cCmb.Name = "Stato";
            //cCmb.Width = 100;
            //cCmb.DataSource = tSta;
            //cCmb.ValueMember = "tab_cod";
            //cCmb.DisplayMember = "tab_des";
            //cCmb.DisplayStyle = DataGridViewComboBoxDisplayStyle.DropDownButton;
            //dgv1.Columns.Add((DataGridViewColumn)cCmb);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "tmp_acq";
            cTbc.Name = "Acquistato";
            cTbc.Width = 60;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = false;
            dgv1.Columns.Add(cTbc);
            dgv1.Columns[cTbc.Name].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dgv1.Columns[cTbc.Name].DefaultCellStyle.Format = "##0.00";

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "tmp_ven";
            cTbc.Name = "Venduto";
            cTbc.Width = 60;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = false;
            dgv1.Columns.Add(cTbc);
            dgv1.Columns[cTbc.Name].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dgv1.Columns[cTbc.Name].DefaultCellStyle.Format = "##0.00";

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "StaDes";
            cTbc.Name = "Stato";
            cTbc.Width = 100;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = true;
            //cTbc.Selected = true;
            dgv1.Columns.Add(cTbc);
            //dgv1.Columns[cTbc.Name].DefaultCellStyle.BackColor = Color.GreenYellow;
        }

        private void FillDati()
        {
            string s = "";
            DataRow x;

            DataTable t = new clsGenTabTmp().TabTmpArtMovimenti("TabArt");
            DataColumn[] keys = new DataColumn[1];
            keys[0] = t.Columns["tmp_art"];
            t.PrimaryKey = keys;

            s = "SELECT ";
            s += "AnaArticoli.art_cod, ";
            s += "AnaArticoli.art_des, ";
            s += "AnaArticoli.art_umi, ";
            s += "AnaArticoli.art_sta, ";
            s += "AnaArticoli.art_rep, ";
            s += "AnaArticoli.art_ec1, ";
            s += "AnaArticoli.art_ec2, ";
            s += "AnaArticoli.art_ec3, ";
            s += "TabStato.tab_des AS StaDes, ";
            s += "TabReparti.tab_des AS RepDes, ";
            s += "TabEcrLv3.tab_des AS Lv3Des ";
            s += "FROM AnaArticoli ";
            s += "LEFT JOIN TabStato ON AnaArticoli.art_sta = TabStato.tab_cod ";
            s += "LEFT JOIN TabReparti ON AnaArticoli.art_rep = TabReparti.tab_cod ";
            s += "LEFT JOIN TabEcrLv3 ON AnaArticoli.art_ec1 = TabEcrLv3.tab_lv1 AND AnaArticoli.art_ec2 = TabEcrLv3.tab_lv2 AND AnaArticoli.art_ec3 = TabEcrLv3.tab_cod ";
            s += "WHERE art_sta='C' OR art_sta='A' OR art_sta='N' ";
            DataTable tArt = _clsFun.FillTabSql("AnaArticoli", s, false, _strConSql);
            keys = new DataColumn[1];
            keys[0] = tArt.Columns["art_cod"];
            tArt.PrimaryKey = keys;

            tArt.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.Decimal"),
                ColumnName = "tmp_acq",
                Caption = "Acquistato",
                ReadOnly = false,
                DefaultValue = (Decimal)0
            });
            tArt.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.Decimal"),
                ColumnName = "tmp_ven",
                Caption = "Venduto",
                ReadOnly = false,
                DefaultValue = (Decimal)0
            });

            //progressBar1.Value = 0;
            //progressBar1.Maximum = tArt.Rows.Count;
            //progressBar1.Minimum = 0;

            //foreach(DataRow y in tArt.Rows)
            //{
            //    progressBar1.Increment(1);
            //    Application.DoEvents();

            //    x = t.NewRow();
            //    x["tmp_ecr"] = (string)y["art_ec1"] + (string)y["art_ec2"] + (string)y["art_ec3"];
            //    x["tmp_ecd"] = y["Lv3Des"];
            //    x["tmp_rep"] = y["art_rep"];
            //    x["tmp_red"] = y["RepDes"];
            //    x["tmp_art"] = y["art_cod"];
            //    x["tmp_ard"] = y["art_des"];
            //    x["tmp_umi"] = y["art_umi"];
            //    x["tmp_acq"] = 0;
            //    x["tmp_ven"] = 0;

            FillMovimenti(tArt);
            
            FillVendite(tArt);

            dgv1.DataSource = tArt;
        }

        private void FillMovimenti(DataTable tabArt)
        {
            string s = "";
            DataRow x;
            DataRow[] j;

            s = "SELECT ";
            s += "GesFatTestate.fat_tpd, ";
            s += "GesMovimenti.mov_art, ";
            s += "SUM(GesMovimenti.mov_qta) AS mov_qta, ";
            s += "SUM(GesMovimenti.mov_qkg) AS mov_qkg ";
            s += "FROM GesMovimenti ";
            s += "INNER JOIN GesFatTestate ON GesMovimenti.mov_yfa = GesFatTestate.fat_yfa AND GesMovimenti.mov_nfa = GesFatTestate.fat_nfa ";
            s += "WHERE ";
            //s += "(GesMovimenti.mov_art = '" + rowTmp["tmp_art"] + "') AND ";
            s += "(GesFatTestate.fat_ddo >= " + _clsFun.DaySql(dtpIni.Value) + " AND GesFatTestate.fat_ddo <= " + _clsFun.DaySql(dtpFin.Value) + ") AND ";
            s += "(GesMovimenti.mov_ori = '') ";
            s += "GROUP BY GesFatTestate.fat_tpd, GesMovimenti.mov_art";
            DataTable tMov = _clsFun.FillTabSql("GesMovimenti", s, false, _strConSql);

            progressBar1.Value = 0;
            progressBar1.Maximum = tMov.Rows.Count;
            progressBar1.Minimum = 0;

            if (tMov.Rows.Count > 0)
            {
                foreach (DataRow y in tMov.Rows)
                {
                    progressBar1.Increment(1);
                    Application.DoEvents();

                    j = tabArt.Select("art_cod='" + y["mov_art"] +"'");
                    if (j.Length > 0)
                    {
                        if ((string)y["fat_tpd"] == "FV")
                        {
                            if ((string)j[0]["art_umi"] == "KG")
                                j[0]["tmp_ven"] = (decimal)j[0]["tmp_ven"] + (decimal)y["mov_qkg"];
                            else
                                j[0]["tmp_ven"] = (decimal)j[0]["tmp_ven"] + (decimal)y["mov_qta"];
                        }
                        else
                        {
                            if ((string)j[0]["art_umi"] == "KG")
                                j[0]["tmp_acq"] = (decimal)j[0]["tmp_acq"] + (decimal)y["mov_qkg"];
                            else
                                j[0]["tmp_acq"] = (decimal)j[0]["tmp_acq"] + (decimal)y["mov_qta"];
                        }
                    }
                }
            }
        }

        private void FillVendite(DataTable tabArt)
        {
            string s = "";
            DataRow x;
            DataRow[] j;

            s = "SELECT ";
            s += "ven_art, ";
            s += "SUM(ven_qta) AS ven_qta, ";
            s += "SUM(ven_qkg) AS ven_qkg ";
            s += "FROM GesNegVen ";
            s += "WHERE (ven_day >= " + _clsFun.DaySql(dtpIni.Value) + " AND ven_day <= " + _clsFun.DaySql(dtpFin.Value) + " ) ";
            s += "GROUP BY ven_art";
            DataTable tVen = _clsFun.FillTabSql("GesVendite", s, false, _strConSqlSta);

            if (tVen.Rows.Count > 0)
            {
                progressBar1.Value = 0;
                progressBar1.Maximum = tVen.Rows.Count;
                progressBar1.Minimum = 0;

                foreach (DataRow y in tVen.Rows)
                {
                    progressBar1.Increment(1);
                    Application.DoEvents();

                    j = tabArt.Select("art_cod='" + (string)y["ven_art"] + "'");
                    if (j.Length > 0)
                    {
                        if ((string)j[0]["art_umi"] == "KG")
                            j[0]["tmp_ven"] = (decimal)j[0]["tmp_ven"] + (decimal)y["ven_qkg"];
                        else
                            j[0]["tmp_ven"] = (decimal)j[0]["tmp_ven"] + (decimal)y["ven_qta"];
                    }
                }
            }

        }

        private void dgv1_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            if(e.ColumnIndex == 3)
            {
                CurrencyManager cm = dgv1.BindingContext[dgv1.DataSource, dgv1.DataMember] as CurrencyManager;
                if (cm.Position >= 0)
                {
                    DataRowView r = cm.List[dgv1.CurrentRow.Index] as DataRowView;
                    DataRow x = r.Row;

                    frmAnaArticolo f = new frmAnaArticolo();
                    f._strArtCod = (string)x["art_cod"];
                    f.ShowDialog();
                }
            }
            else if(e.ColumnIndex == 7)
            {

                CurrencyManager cm = dgv1.BindingContext[dgv1.DataSource, dgv1.DataMember] as CurrencyManager;
                if (cm.Position >= 0)
                {
                    DataRowView r = cm.List[dgv1.CurrentRow.Index] as DataRowView;
                    DataRow x = r.Row;

                    string s = (string)x["art_sta"];
                    if (s == "C")
                    {
                        x["art_sta"] = "N";
                        x["StaDes"] = "NON ATTIVO";
                    }
                    else if (s == "N")
                    {
                        x["art_sta"] = "A";
                        x["StaDes"] = "ATTIVO";
                    }
                    else if (s == "A")
                    {
                        x["art_sta"] = "C";
                        x["StaDes"] = "CANCELLATO";
                    }

                    s = "UPDATE AnaArticoli SET art_sta='" + x["art_sta"] + "' WHERE art_cod='" + x["art_cod"] + "'";
                    _clsFun.SqlWrite(s, _strConSql);
                    _clsVar.Variazioni((string)x["art_cod"], s, "Movimenti articolo", _clsDef.VARPOS);
                }

            }
        }

        private void dgv1_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            if (dgv1.Columns[e.ColumnIndex].Name == "Stato")
            {
                e.CellStyle.BackColor = Color.Aquamarine;
                e.CellStyle.BackColor = Color.Black;

                string s = (string)dgv1.Rows[e.RowIndex].Cells[e.ColumnIndex].Value;
                if (s.Length > 0)
                {
                    if (s.Substring(0, 1) == _clsDef.STAATT)
                        e.CellStyle.BackColor = Color.PaleGreen;
                    else if (s.Substring(0, 1) == _clsDef.STANOA)
                        e.CellStyle.BackColor = Color.Crimson;
                    else if (s.Substring(0, 1) == _clsDef.STACAN)
                    {
                        e.CellStyle.BackColor = Color.Black;
                        e.CellStyle.ForeColor = Color.White;
                    }
                }
            }
        }

    }
}
