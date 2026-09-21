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
    public partial class frmGesStatCostiAggiorna : Form
    {
        clsDefine _clsDef = new clsDefine();
        clsFuncs _clsFun = new clsFuncs();

        private string _strConSql = "";
        private string _strConSqlSta = "";
        private string _strForPri = "";

        public frmGesStatCostiAggiorna()
        {
            InitializeComponent(); 
            new clsGesGraph().SetGraph(this, 0);
        }

        private void frmGesStatCostiAggiorna_Load(object sender, EventArgs e)
        {
            _strConSql = _clsFun.ConSql("");
            _strConSqlSta = _clsFun.ConSql("3");

            _strForPri = _clsFun.ParGet(clsDefine.enuParametri.ParForDefault, _strConSql);
            string[] a = _strForPri.Split(',');
            if(a.Length > 0)
                _strForPri = a[0];

            dtpIni.Value = DateTime.Today.AddDays(-15);
            SetDgv1();
        }

        private void esciToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Esci();
        }

        private void frmGesStatCostiAggiorna_KeyDown(object sender, KeyEventArgs e)
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
            dgv1.AutoGenerateColumns = false;
            //dgv1.VirtualMode = true;
            //dgv1.Dock = DockStyle.Fill;
            dgv1.AllowUserToAddRows = false;
            dgv1.ReadOnly = false;
            dgv1.AllowUserToDeleteRows = false;
            //dgv1.DisplayedRowCount() = true;

            DataGridViewTextBoxColumn cTbc;
            //DataGridViewCheckBoxColumn cCbc;
            //DataGridViewComboBoxColumn cCmb;

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "ven_day";
            cTbc.Name = "Data";
            cTbc.Width = 80;
            cTbc.ValueType = typeof(string);
            cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            cTbc.ReadOnly = true;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "ven_art";
            cTbc.Name = "Articolo";
            cTbc.Width = 60;
            cTbc.ValueType = typeof(string);
            //cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            cTbc.ReadOnly = true;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "ven_ard";
            cTbc.Name = "Descrizione";
            cTbc.Width = 250;
            cTbc.ValueType = typeof(string);
            //cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            cTbc.ReadOnly = true;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "ven_cos";
            cTbc.Name = "Costo";
            cTbc.Width = 80;
            cTbc.ValueType = typeof(string);
            cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            cTbc.DefaultCellStyle.Format = "###,##0.00";
            cTbc.ReadOnly = false;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "CosTip";
            cTbc.Name = "Tipo costo";
            cTbc.Width = 25;
            cTbc.ValueType = typeof(string);
            //cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            cTbc.ReadOnly = true;
            dgv1.Columns.Add(cTbc);
        }

        private void FillDati()
        {
            string s = "";
            DataRow[] j;

            s = "SELECT DISTINCT ";
            s += "lia_tip, ";
            s += "lia_art, ";
            s += "lia_arf, ";
            s += "lia_for, ";
            s += "lia_dti, ";
            s += "lia_dtf, ";
            s += "lia_cos ";
            s += "FROM GesLisAcquisto ";
            s += "WHERE (lia_ann = 0) ";
            s += "ORDER BY lia_art, lia_tip";
            DataTable tLia = _clsFun.FillTabSql("TabLia", s, false, _strConSql);
            DataColumn[] keys = new DataColumn[6];
            keys[0] = tLia.Columns["lia_tip"];
            keys[1] = tLia.Columns["lia_art"];
            keys[2] = tLia.Columns["lia_dti"];
            keys[3] = tLia.Columns["lia_for"];
            keys[4] = tLia.Columns["lia_arf"];
            keys[5] = tLia.Columns["lia_dtf"];
            tLia.PrimaryKey = keys;

            s = "SELECT ";
            s += "ven_idx, ven_day, ven_art, ven_ard, ven_cos ";
            s += "FROM GesNegVen ";
            s += "WHERE (ven_day >= " + _clsFun.DaySql(dtpIni.Value) + ") AND (ven_day <= " + _clsFun.DaySql(dtpFin.Value) + ")";
            if (!chkAll.Checked)
                s += " AND (ven_cos IS NULL) OR ven_cos = 0 ";
            DataTable tSta = _clsFun.FillTabSql("TabSta", s, false, _strConSqlSta);

            tSta.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "CosTip",
                Caption = "Tip",
                MaxLength = 1,
                ReadOnly = false,
                DefaultValue = (String)""
            });

            progressBar1.Value = 0;
            progressBar1.Maximum = tSta.Rows.Count;
            progressBar1.Minimum = 0;

            foreach (DataRow y in tSta.Rows)
            {
                progressBar1.Increment(1);
                System.Windows.Forms.Application.DoEvents();

                j = tLia.Select("lia_tip='M' AND lia_art='" + (string)y["ven_art"] + "' AND lia_dti <= " + _clsFun.DayMdb((DateTime)y["ven_day"]), "lia_dti DESC");
                if(j.Length == 0)
                    j = tLia.Select("lia_tip='F' AND lia_art='" + (string)y["ven_art"] + "' AND lia_dti <= " + _clsFun.DayMdb((DateTime)y["ven_day"]), "lia_dti DESC");

                if (j.Length == 0)
                {
                    j = tLia.Select("lia_tip='O' AND lia_art='" + (string)y["ven_art"] + "' AND lia_dti <= " + _clsFun.DayMdb((DateTime)y["ven_day"]), "lia_dti DESC");

                    if (j.Length == 0)
                    {
                        j = tLia.Select("lia_tip='L' AND lia_art='" + (string)y["ven_art"] + "' AND lia_dti <= " + _clsFun.DayMdb((DateTime)y["ven_day"]) + " AND lia_for='" + _strForPri + "'", "lia_dti DESC");
                        if (j.Length == 0)
                            j = tLia.Select("lia_tip='L' AND lia_art='" + (string)y["ven_art"] + "' AND lia_dti <= " + _clsFun.DayMdb((DateTime)y["ven_day"]), "lia_dti DESC");
                    }
                }
                if (j.Length > 0)
                {
                    y["ven_cos"] = j[0]["lia_cos"];
                    y["CosTip"] = j[0]["lia_tip"];

                    s = "UPDATE GesNegVen SET ven_cos=" + ((decimal)y["ven_cos"]).ToString().Replace(",", ".") + " WHERE ";
                    s += "ven_idx =" + Convert.ToString(y["ven_idx"]);

                    _clsFun.SqlWrite(s, _strConSqlSta);
                }
            }

            dgv1.DataSource = tSta;
        }

        private void dgv1_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            if (dgv1.Columns[e.ColumnIndex].Name == "Descrizione")
            {
                CurrencyManager cm = dgv1.BindingContext[dgv1.DataSource, dgv1.DataMember] as CurrencyManager;
                if (cm.Position >= 0)
                {
                    DataRowView r = cm.List[dgv1.CurrentRow.Index] as DataRowView;
                    DataRow x = r.Row;

                    string s = ((string)x["ven_art"]).Trim();

                    if (s != "")
                    {
                        frmAnaArticolo f = new frmAnaArticolo();
                        f._strArtCod = s;
                        f.ShowDialog();

                        //if (f._strSqlArt != "" && f._tabArt.Rows.Count > 0)
                        //{
                        //    if (f._strSqlArt.Contains("art_rep"))
                        //    {
                        //        DataTable t = f._tabArt;
                        //        if ((string)x["div_rep"] != (string)t.Rows[0]["art_rep"] && (string)t.Rows[0]["art_rep"] != _strRepDef && _strRepDef != "")
                        //            x["div_rep"] = (string)t.Rows[0]["art_rep"];
                        //    }
                        //}
                    }
                }
            }

        }

    }
}
