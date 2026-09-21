using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace APOffice
{
    public partial class frmSeekEcrArt : Form
    {
        private const string TABLIV = "TabEcrLv";
        private const string TABLV1 = "TabEcrLv1";
        private const string TABLV2 = "TabEcrLv2";
        private const string TABLV3 = "TabEcrLv3";
        private const string TABANAART = "AnaAericoli";

        clsDefine _clsDef = new clsDefine();
        clsFuncs _clsFun = new clsFuncs();
        clsQuery _clsQry = new clsQuery();

        private DataSet _dasGen = new DataSet();
        private string _strConSql = "";

        public string _strArtCod = "";

        public frmSeekEcrArt()
        {
            InitializeComponent();
            new clsGesGraph().SetGraph(this, 0);;
        }

        private void frmGesTabEcr_Load(object sender, EventArgs e)
        {
            _strConSql = _clsFun.ConSql("");
            SetDgv(dgv1);
            SetDgv(dgv2);
            SetDgv(dgv3);
            SetDgvArt(dgv4);
            FillDati();
            txtSeek.Select();
        }
        private void frmSeekEcrArt_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Escape)
                Esci();
        }
        private void esciToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //Salva();
            Esci();
        }
        private void Esci()
        {
            this.Close();
        }

        private void SetDgv(DataGridView dgv)
        {
            dgv.AutoGenerateColumns = false;
            //dgv.VirtualMode = true;
            //dgv.Dock = DockStyle.Fill;
            dgv.AllowUserToAddRows = false;
            dgv.ReadOnly = false;
            dgv.AllowUserToDeleteRows = false;
            //dgv.DisplayedRowCount() = true;

            DataGridViewTextBoxColumn cTbc;
            DataGridViewCheckBoxColumn cCbc;
            //DataGridViewComboBoxColumn cCmb;

            //cTbc = new DataGridViewTextBoxColumn();
            //cTbc.DataPropertyName = "tab_cod";
            //cTbc.Name = "Codice";
            //cTbc.Width = 30;
            //if (dgv.Name == "dgv3")
            //    cTbc.Width = 50;
            //cTbc.ValueType = typeof(string);
            //dgv.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "tab_des";
            cTbc.Name = "Descrizione";
            cTbc.Width = 300;
            cTbc.ValueType = typeof(string);
            dgv.Columns.Add(cTbc);

            //cCbc = new DataGridViewCheckBoxColumn();
            //cCbc.ValueType = typeof(Boolean);
            //cCbc.DataPropertyName = "tab_ann";
            //cCbc.Name = "Ann.";
            //cCbc.Width = 40;
            //dgv.Columns.Add(cCbc);
        }

        private void SetDgvArt(DataGridView dgv)
        {
            dgv.AutoGenerateColumns = false;
            //dgv.VirtualMode = true;
            //dgv.Dock = DockStyle.Fill;
            dgv.AllowUserToAddRows = false;
            dgv.ReadOnly = false;
            dgv.AllowUserToDeleteRows = false;
            //dgv.DisplayedRowCount() = true;

            DataGridViewTextBoxColumn cTbc;
            DataGridViewCheckBoxColumn cCbc;
            //DataGridViewComboBoxColumn cCmb;

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "art_cod";
            cTbc.Name = "Codice";
            cTbc.Width = 150;
            cTbc.ValueType = typeof(string);
            dgv.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "art_des";
            cTbc.Name = "Descrizione";
            cTbc.Width = 600;
            cTbc.ValueType = typeof(string);
            dgv.Columns.Add(cTbc);

            //cCbc = new DataGridViewCheckBoxColumn();
            //cCbc.ValueType = typeof(Boolean);
            //cCbc.DataPropertyName = "tab_ann";
            //cCbc.Name = "Ann.";
            //cCbc.Width = 40;
            //dgv.Columns.Add(cCbc);
        }

        private void FillDati()
        {
            DataColumn c;
            string p = "";
            string s = "";
            DataTable t;

            for (int i = 1; i <= 3; i++)
            {
                p = "TabEcrLv" + i.ToString();
                s = "SELECT * FROM " + p + " ORDER BY tab_cod";
                t = _clsFun.FillTabSql(p, s, false, _strConSql);

                c = new DataColumn();
                c.DataType = System.Type.GetType("System.String");
                c.ColumnName = "TabMdy";
                c.Caption = "Modify";
                c.MaxLength = 1;
                c.ReadOnly = false;
                c.DefaultValue = "";
                t.Columns.Add(c);

                _dasGen.Tables.Add(t);
            }

            dgv1.DataSource = new DataView(_dasGen.Tables[TABLV1], "", "", DataViewRowState.CurrentRows);
        }

        //private void btnNew_Click(object sender, EventArgs e)
        //{
        //    Button btn = (Button)sender;
        //    string sTab = TABLIV + btn.Name.Substring(6, 1);
        //    if (sTab == TABLV1)
        //        NewRec(sTab, dgv1);
        //    if (sTab == TABLV2)
        //        NewRec(sTab, dgv2);
        //    if (sTab == TABLV3)
        //        NewRec(sTab, dgv3);
        //    //if (sTab == TABLV4)
        //    //    GesCfo("NEW");
        //}

        private void OldNewRec(string strTab, DataGridView dgv)
        {
            DataRow y = _dasGen.Tables[strTab].NewRow();
            y["tab_cod"] = "";
            y["tab_des"] = "";
            y["tab_ann"] = false;

            if (strTab == TABLV1)
            {
                y["tab_doa"] = "";
                y["tab_bil"] = "";
            }
            if (strTab == TABLV2)
            {
                CurrencyManager cm = dgv1.BindingContext[dgv1.DataSource, dgv1.DataMember] as CurrencyManager;
                DataRowView r = cm.List[cm.Position] as DataRowView;
                y["tab_lv1"] = r.Row["tab_cod"];
            }
            if (strTab == TABLV3)
            {
                CurrencyManager cm = dgv2.BindingContext[dgv2.DataSource, dgv2.DataMember] as CurrencyManager;
                DataRowView r = cm.List[cm.Position] as DataRowView;
                y["tab_lv1"] = r.Row["tab_lv1"];
                y["tab_lv2"] = r.Row["tab_cod"];
            }

            _dasGen.Tables[strTab].Rows.Add(y);

            dgv.ClearSelection();

            int i = 0;

            foreach (DataGridViewRow r in dgv.Rows)
            {

                if (r.Cells[0].Value.ToString() == (""))
                {
                    dgv.CurrentCell = r.Cells[0];
                    dgv.Rows[i].Selected = true;
                    //GesRow(dgv);
                    i++;
                    //Console.WriteLine("aaaa");
                    break;
                }
            }
        }

        //private void NewRec(string strTab, DataGridView dgv)
        //{
        //    int i = 0;
        //    //string s = "";
        //    string sWhe = "";
        //    string sLv1 = "";
        //    string sLv2 = "";
        //    if (dgv.Name == "dgv1")
        //    {
        //        //CurrencyManager cm = dgv1.BindingContext[dgv1.DataSource, dgv1.DataMember] as CurrencyManager;
        //        //DataRowView r = cm.List[cm.Position] as DataRowView;
        //        //sLv1 = (string)r.Row["tab_cod"];
        //        //sWhe = "tab_des='' AND tab_lv1='" + sLv1 + "'";
        //        sWhe = "tab_des=''";
        //    }
        //    if (dgv.Name == "dgv2")
        //    {
        //        CurrencyManager cm = dgv1.BindingContext[dgv1.DataSource, dgv1.DataMember] as CurrencyManager;
        //        DataRowView r = cm.List[cm.Position] as DataRowView;
        //        sLv1 = (string)r.Row["tab_cod"];
        //        sWhe = "tab_des='' AND tab_lv1='" + sLv1 + "'";
        //    }
        //    if (dgv.Name == "dgv3")
        //    {
        //        CurrencyManager cm = dgv2.BindingContext[dgv2.DataSource, dgv2.DataMember] as CurrencyManager;
        //        DataRowView r = cm.List[cm.Position] as DataRowView;
        //        sLv1 = (string)r.Row["tab_lv1"];
        //        sLv2 = (string)r.Row["tab_cod"];
        //        sWhe = "tab_des='' AND tab_lv1='" + sLv1 + "' AND tab_lv1='" + sLv2 + "'";
        //    }
            
        //    DataView v = new DataView(_dasGen.Tables[strTab], sWhe, "", DataViewRowState.CurrentRows);
        //    if (v.Count > 0)
        //    {
        //        MessageBox.Show("Ci sono righe senza descrizione!");
        //        i = 0;

        //        foreach (DataGridViewRow r in dgv.Rows)
        //        {
        //            if (r.Cells[0].Value.ToString() == ((v[0]["tab_cod"]).ToString()))
        //            {
        //                dgv.CurrentCell = r.Cells[0];
        //                dgv.Rows[i].Selected = true;
        //                i++;
        //                break;
        //            }
        //        }

        //        dgv.CurrentCell = dgv.Rows[dgv.RowCount - 1].Cells[1];
        //        dgv.Rows[dgv.RowCount - 1].Selected = true;
        //        dgv.BeginEdit(true);
        //    }
        //    else
        //    {
        //        string sCod = "";
        //        sWhe = "";
        //        DataRow y = _dasGen.Tables[strTab].NewRow();
        //        y["tab_cod"] = "";
        //        y["tab_des"] = "";
        //        y["tab_ann"] = false;
        //        y["TabMdy"] = "S";
        //        if (dgv.Name == "dgv2")
        //        {
        //            y["tab_lv1"] = sLv1;
        //            sWhe = "tab_lv1='" + sLv1 + "'";
        //        }
        //        if (dgv.Name == "dgv3")
        //        {
        //            y["tab_lv1"] = sLv1;
        //            y["tab_lv2"] = sLv2;
        //            sWhe = "tab_lv1='" + sLv1 + "' AND tab_lv2='" + sLv2 + "'";
        //        }
 
        //        sCod = "000";
        //        v = new DataView(_dasGen.Tables[strTab], sWhe, "tab_cod DESC", DataViewRowState.CurrentRows);
        //        if (v.Count > 0)
        //            sCod = (string)v[0]["tab_cod"];
        //        sCod = (Convert.ToInt16(sCod) + 1).ToString("000");
        //        y["tab_cod"] = sCod;

        //        _dasGen.Tables[strTab].Rows.Add(y);

        //        if (dgv.Name == "dgv1")
        //            dgv.DataSource = _dasGen.Tables[TABLV1];
        //        else
        //        {
        //            if (dgv.Name == "dgv2")
        //                v = new DataView(_dasGen.Tables[TABLV2], sWhe, "tab_cod", DataViewRowState.CurrentRows);
        //            else if (dgv.Name == "dgv3")
        //                v = new DataView(_dasGen.Tables[TABLV3], sWhe, "tab_cod", DataViewRowState.CurrentRows);
        //            dgv.DataSource = v;
        //        }

        //        dgv.ClearSelection();

        //        i = 0;

        //        foreach (DataGridViewRow r in dgv.Rows)
        //        {
        //            if (r.Cells[0].Value.ToString() == (sCod))
        //            {
        //                dgv.CurrentCell = r.Cells[0];
        //                dgv.Rows[i].Selected = true;
        //                i++;
        //            }
        //        }

        //        dgv.CurrentCell = dgv.Rows[dgv.RowCount - 1].Cells[1];
        //        dgv.Rows[dgv.RowCount - 1].Selected = true;
        //        dgv.BeginEdit(true);
        //    }
        //}

        private void dgv1_RowEnter(object sender, DataGridViewCellEventArgs e)
        {
            CurrencyManager cm = this.dgv1.BindingContext[this.dgv1.DataSource, this.dgv1.DataMember] as CurrencyManager;

            if (cm.Position >= 0)
            {

                DataRowView r = cm.List[e.RowIndex] as DataRowView;
                string sCod = (string)r.Row["tab_cod"];

                DataView v = new DataView(_dasGen.Tables[TABLV2], "tab_lv1='" + sCod + "'", "tab_cod", DataViewRowState.CurrentRows);
                dgv2.DataSource = v;

                dgv3.DataSource = null;
            }
        }

        private void dgv2_RowEnter(object sender, DataGridViewCellEventArgs e)
        {
            CurrencyManager cm = this.dgv2.BindingContext[this.dgv2.DataSource, this.dgv2.DataMember] as CurrencyManager;

            if (cm.Position >= 0)
            {

                DataRowView r = cm.List[e.RowIndex] as DataRowView;
                string sCod = Convert.ToString(r.Row["tab_cod"]);
                string sLv1 = Convert.ToString(r.Row["tab_lv1"]);

                DataView v = new DataView(_dasGen.Tables[TABLV3], "tab_lv1='" + sLv1 + "' AND tab_lv2='" + sCod + "'", "tab_cod", DataViewRowState.CurrentRows);

                dgv3.DataSource = v;
            }
        }

        private void dgv3_RowEnter(object sender, DataGridViewCellEventArgs e)
        {
            CurrencyManager cm = dgv3.BindingContext[dgv3.DataSource, dgv3.DataMember] as CurrencyManager;

            if (cm.Position >= 0)
            {
                DataRowView r = cm.List[e.RowIndex] as DataRowView;
                string sCod = Convert.ToString(r.Row["tab_cod"]);
                string sLv2 = Convert.ToString(r.Row["tab_lv2"]);
                string sLv1 = Convert.ToString(r.Row["tab_lv1"]);

                string s = "SELECT art_cod, art_des FROM AnaArticoli WHERE art_ec1='" + sLv1 + "' AND art_ec2='" + sLv2 + "' AND art_ec3='" + sCod + "' ORDER BY art_des";

                DataTable t = _clsFun.FillTabSql(TABANAART, s, false, _strConSql);

                dgv4.DataSource = t;


                //DataView v = new DataView(_dasGen.Tables[TABLV4], "int_con='" + sLv1 + sLv2 + sCod + "'", "int_con", DataViewRowState.CurrentRows);
                //dgv4.DataSource = v;
            }
        }



        private void txtSeek_KeyDown(object sender, KeyEventArgs e)
        {
            if(e.KeyCode == Keys.Enter)
                SeekDes();
        }

        private void btnOK_Click(object sender, EventArgs e)
        {
            SeekDes();
        }

        private void SeekDes()
        {
            if (txtSeek.Text != "")
            {
                string s = "";
                string sWhe = "";

                if (txtSeek.Text.Contains("+"))
                {
                    s = txtSeek.Text;

                    string[] a = s.Split('+');
                    s = "";
                    for (int i = 0; i < a.Length; i++)
                    {
                        s += "art_des LIKE '%" + a[i] + "%' AND ";
                    }
                    sWhe = s.Substring(0, s.Length - 4);
                }
                else
                    sWhe = "art_des LIKE '%" + txtSeek.Text.Trim() + "%'";

                //s = "SELECT art_cod, art_des FROM AnaArticoli WHERE  art_des LIKE '%" + s + "%' ORDER BY art_des";
                s = "SELECT art_cod, art_des FROM AnaArticoli WHERE " + sWhe;

                DataTable t = _clsFun.FillTabSql(TABANAART, s, false, _strConSql);
                dgv4.DataSource = t;

                if (t.Rows.Count > 0)
                    dgv4.Select();

            }
        }

        private void dgv4_RowEnter(object sender, DataGridViewCellEventArgs e)
        {
            //CurrencyManager cm = dgv4.BindingContext[dgv4.DataSource, dgv4.DataMember] as CurrencyManager;
            //if (cm.Position >= 0)
            //{
            //    DataRowView r = cm.List[dgv4.CurrentRow.Index] as DataRowView;
            //    DataRow x = r.Row;
            //    _strArtCod = (string)x["art_cod"];
            //}
        }

        private void dgv4_Click(object sender, EventArgs e)
        {
            Scelto();
        }

        private void Scelto()
        {
            CurrencyManager cm = dgv4.BindingContext[dgv4.DataSource, dgv4.DataMember] as CurrencyManager;
            if (cm.Position >= 0)
            {
                DataRowView r = cm.List[dgv4.CurrentRow.Index] as DataRowView;
                DataRow x = r.Row;
                _strArtCod = (string)x["art_cod"];
                Esci();
            }
        }

        private void dgv_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            DataGridView dgv = (DataGridView)sender;

            CurrencyManager cm = dgv.BindingContext[dgv.DataSource, dgv.DataMember] as CurrencyManager;
            if (cm.Position >= 0)
            {
                DataRowView r = cm.List[dgv.CurrentRow.Index] as DataRowView;
                DataRow x = r.Row;
                x["TabMdy"] = "S";
            }
        }

        private void lblFreccia_Click(object sender, EventArgs e)
        {
            Label lbl = (Label)sender;
            GriPos(lbl.Name, "1_CLICK");
        }

        private void lblFreccia_DoubleClick(object sender, EventArgs e)
        {
            Label lbl = (Label)sender;
            GriPos(lbl.Name, "2_CLICK");
        }

        private void GriPos(string strCmd, string strClick)
        {
            if (dgv4.DataSource != null && dgv4.Rows.Count > 0)
            {
                dgv4.ClearSelection();

                if (strCmd == "ROWLAST" || (strCmd == "lblGriGiu" && strClick == "2_CLICK"))
                {
                    int nRowIndex = dgv4.Rows.Count - 1;
                    int nColumnIndex = 1;

                    dgv4.FirstDisplayedScrollingRowIndex = nRowIndex;

                    dgv4.Rows[nRowIndex].Selected = true;
                    dgv4.Rows[nRowIndex].Cells[1].Selected = true;
                    dgv4.CurrentCell = dgv4[nColumnIndex, nRowIndex];
                }
                else if (strCmd == "ROWFIRST" || (strCmd == "lblGriSu" && strClick == "2_CLICK"))
                {
                    int nRowIndex = 0;
                    int nColumnIndex = 1;

                    dgv4.FirstDisplayedScrollingRowIndex = nRowIndex;

                    dgv4.Rows[nRowIndex].Selected = true;
                    dgv4.Rows[nRowIndex].Cells[1].Selected = true;
                    dgv4.CurrentCell = dgv4[nColumnIndex, nRowIndex];
                }
                else if (strCmd.Substring(0, 3) == "lbl" && strClick == "1_CLICK")
                {
                    string sCmd = strCmd.Substring(strCmd.Length - 3, 3);

                    int nRowIndex = dgv4.CurrentRow.Index;
                    if (sCmd == "iSu" && nRowIndex > 0)
                        nRowIndex = nRowIndex - 1;
                    else if (sCmd == "Giu" == nRowIndex < dgv4.Rows.Count - 1)
                        nRowIndex = nRowIndex + 1;

                    dgv4.Rows[nRowIndex].Cells[1].Selected = true;
                    dgv4.CurrentCell = dgv4[1, nRowIndex];
                }
            }
        }

        private void Salva()
        {
            DataRow[] j;
            string sSql = "";

            for (int i = 1; i <= 3; i++)
            {
                string p = TABLIV + i.ToString();
                string s = "SELECT * FROM " + p;
                DataTable t = _clsFun.FillTabSql(p, s, false, _strConSql);

                DataTable tTmp = t.Clone();

                foreach (DataRow y in _dasGen.Tables[p].Rows)
                {
                    if ((string)y["TabMdy"] != "")
                    {
                        y["tab_des"] = (string)y["tab_des"];
                        sSql = "tab_cod='" + y["tab_cod"] + "' AND ";
                        for (int i2 = 1; i2 < i; i2++)
                            sSql += "tab_lv" + i2.ToString() + "='" + y["tab_lv" + i2.ToString()] + "' AND ";

                        sSql = sSql.Substring(0, sSql.Length - 5);
                        j = t.Select(sSql);
                        s = "";
                        if (j.Length == 0)
                            s = _clsFun.SqlInsertRow(p, t, y);
                        else
                        {
                            ArrayList aWhe = new ArrayList();
                            aWhe.Add("tab_cod");
                            for (int i2 = 1; i2 < i; i2++)
                                aWhe.Add("tab_lv" + i2.ToString());
                            ArrayList aExl = new ArrayList();

                            s = _clsFun.SqlUpdRow(p, t, j[0], y, aWhe, aExl);
                        }
                        if (s != "")
                        {
                            _clsFun.SqlWrite(s, _strConSql);
                            tTmp.ImportRow(y);
                        }
                    }
                }

                s = _clsFun.ParGet(clsDefine.enuParametri.Par016PathDivNegozi, _strConSql);
                if (tTmp != null && tTmp.Rows.Count > 0 && s != "")
                    new clsVariazioni().DivNegTabelle(s, tTmp);
            }
        }

        private void btnSeek_Click(object sender, EventArgs e)
        {
            frmSeekArt f = new frmSeekArt();
            f.ShowDialog();
            if (f._tabArt != null && f._tabArt.Rows.Count > 0)
            {
                _strArtCod = (string)f._tabArt.Rows[0]["tmp_art"];
                Esci();
            }
        }

        private void dgv4_KeyDown(object sender, KeyEventArgs e)
        {
            if(e.KeyCode == Keys.Return)
            {
                Scelto();
            }
        }

    }
}
