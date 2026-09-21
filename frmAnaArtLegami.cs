using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace APOffice
{
    public partial class frmAnaArtLegami : Form
    {
        clsDefine _clsDef = new clsDefine();
        clsFuncs _clsFun = new clsFuncs();
        clsQuery _clsQry = new clsQuery();
        clsVariazioni _clsVar = new clsVariazioni();

        public string _strArt = "";
        public string _strArd = "";
        public string _strTip = "";
        public string _strKey = "";

        private string _strConSql = "";

        public frmAnaArtLegami()
        {
            InitializeComponent();
            new clsGesGraph().SetGraph(this, 0);
            _strConSql = _clsFun.ConSql("");
        }

        private void frmAnaArtDistinta_Load(object sender, EventArgs e)
        {
            if (_strTip == "LOT")
            {
                this.Text += " LOTTO";
                lblArt.Text += " - LOTTO " + _strKey;

                if (_strKey.Trim().Length > 3)
                    MessageBox.Show("Lunchezza codice lotto maggiore di 3!", "GENERAZIONE BARCODE", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                else
                    _strKey = _strKey.PadLeft(3, Convert.ToChar("0"));
            }

            SetDgv1();
            FillDati();
            txtSeek.Select();
        }

        private void esciToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Esci();
        }

        private void frmAnaArtDistinta_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Escape)
                Esci();
        }

        private void Esci()
        {
            Salva();
            this.Close();
        }

        private void btnSeek_Click(object sender, EventArgs e)
        {
            frmSeekArt f = new frmSeekArt();
            f.ShowDialog();
            if (f._tabArt != null && f._tabArt.Rows.Count > 0)
            {
                LegArtNew(f._tabArt, 0);
            }
        }

        private void txtSeek_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Return)
            {
                if (txtSeek.Text != "")
                {
                    DataTable t = _clsQry.ArtSeek(txtSeek.Text, "SEEK");
                    if (t != null && t.Rows.Count > 0)
                    {
                        LegArtNew(t, 0);
                    }
                }
            }
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

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "leg_rig";
            cTbc.Name = "Riga";
            cTbc.Width = 20;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = true;
            //cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "leg_leg";
            cTbc.Name = "Articolo";
            cTbc.Width = 60;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = false;
            //cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "art_des";
            cTbc.Name = "Descrizione";
            cTbc.Width = 280;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = true;
            cTbc.MaxInputLength = 90;
            //cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "art_tra";
            cTbc.Name = "Lotto";
            cTbc.Width = 50;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = true;
            cTbc.MaxInputLength = 90;
            cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            cTbc.ToolTipText = "Impostazione lotto in anagrafica articoli";
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "leg_inc";
            cTbc.Name = "Inc. %";
            cTbc.Width = 80;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = false;
            cTbc.MaxInputLength = 90;
            cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dgv1.Columns.Add(cTbc);

            cCbc = new DataGridViewCheckBoxColumn();
            cCbc.DataPropertyName = "leg_ann";
            cCbc.Name = "Ann.";
            cCbc.Width = 30;
            cCbc.ValueType = typeof(string);
            cCbc.ReadOnly = false;
            //cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dgv1.Columns.Add(cCbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "LegMdy";
            cTbc.Name = "Modificato";
            cTbc.Width = 0;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = false;
            cTbc.Visible = true;
            dgv1.Columns.Add(cTbc);
        }

        private void LegArtNew(DataTable tabArt, decimal decPro)
        {
            if (tabArt.Rows.Count > 0)
            {
                DataTable t = (DataTable)dgv1.DataSource;

                DataRow[] j = t.Select("leg_leg='" + tabArt.Rows[0]["tmp_art"] + "'");
                if (j.Length > 0)
                    MessageBox.Show("Articolo già presente!", "CONTROLLO INSERIMENTO", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                else
                {
                    string sRig = "01";
                    for (int i = 1; i <= 99; i++)
                    {
                        sRig = i.ToString("00");
                        j = t.Select("leg_rig='" + sRig + "'");
                        if (j.Length == 0)
                            break;
                    }

                    if(Convert.ToInt16(sRig) > 99)
                        MessageBox.Show("Numero massimo righe 99!", "CONTROLLO INSERIMENTO", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    else
                    {
                        DataRow x = t.NewRow();
                        x["leg_rig"] = sRig;
                        x["leg_art"] = _strArt;
                        x["leg_tip"] = _strTip;
                        x["leg_leg"] = (string)tabArt.Rows[0]["tmp_art"];
                        x["leg_inc"] = 0;
                        x["leg_ann"] = false;
                        x["art_des"] = (string)tabArt.Rows[0]["tmp_ard"];
                        x["LegMdy"] = "I";
                        //((DataView)dgv1.DataSource).Table.Rows.Add(x);
                        t.Rows.Add(x);

                        dgv1.CurrentCell = dgv1.Rows[dgv1.RowCount - 1].Cells[1];
                        dgv1.Rows[dgv1.RowCount - 1].Selected = true;
                    }
                }
            }
        }

        private void FillDati()
        {
            string s = "";

            if (_strArd == "")
            {
                s = "SELECT art_des FROM AnaArticoli WHERE art_cod='" + _strArt + "'";
                DataTable tt = _clsFun.FillTabSql("", s, false, _strConSql);
                if(tt.Rows.Count > 0)
                    _strArd = (string)tt.Rows[0]["art_des"];
            }

            lblArt.Text = _strArt + " " + _strArd;

            s = "SELECT AnaArtLegami.*, AnaArticoli.art_des, AnaArticoli.art_tra FROM AnaArtLegami ";
            s += "LEFT OUTER JOIN AnaArticoli ON AnaArticoli.art_cod = AnaArtLegami.leg_leg ";
            s += "WHERE leg_tip='" + _strTip + "' AND leg_art='" + _strArt + "' ";
            s += "ORDER BY leg_rig";
            DataTable t = _clsFun.FillTabSql("", s, false, _strConSql);

            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "LegMdy",
                Caption = "Mdy",
                MaxLength = 1,
                ReadOnly = false,
                DefaultValue = (String)""
            });

            dgv1.DataSource = t;
            Totali();
        }
        
        private void dgv1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            if (dgv1.CurrentCell is DataGridViewCheckBoxCell)
            {
                dgv1.Rows[e.RowIndex].Cells["Modificato"].Value = "S";
                Totali();
            }
        }

        private void dgv1_CellContentDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            //if (e.ColumnIndex == 2)
            //{
            //    CurrencyManager cm = dgv1.BindingContext[dgv1.DataSource, dgv1.DataMember] as CurrencyManager;
            //    if (cm.Position >= 0)
            //    {
            //        DataRowView r = cm.List[dgv1.CurrentRow.Index] as DataRowView;
            //        DataRow x = r.Row;

            //        string sArt = (string)x["leg_leg"];

            //        frmAnaArticolo f = new frmAnaArticolo();
            //        f._strArtCod = sArt;
            //        f.ShowDialog();
            //        if (f._tabArt != null && f._tabArt.Rows.Count > 0 && (string)f._tabArt.Rows[0]["art_cod"] != sArt)
            //        {
            //            x["leg_leg"] = (string)f._tabArt.Rows[0]["art_cod"];
            //            x["art_des"] = (string)f._tabArt.Rows[0]["art_des"];
            //            x["art_tra"] = (string)f._tabArt.Rows[0]["art_tra"];
            //        }
            //        else if(f._tabArt != null && f._tabArt.Rows.Count > 0)
            //            x["art_tra"] = (string)f._tabArt.Rows[0]["art_tra"];
            //        x["LegMdy"] = "S";
            //    }
            //}
        }
        private void dgv1_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.ColumnIndex == 2)
            {
                CurrencyManager cm = dgv1.BindingContext[dgv1.DataSource, dgv1.DataMember] as CurrencyManager;
                if (cm.Position >= 0)
                {
                    DataRowView r = cm.List[dgv1.CurrentRow.Index] as DataRowView;
                    DataRow x = r.Row;

                    string sArt = (string)x["leg_leg"];

                    frmAnaArticolo f = new frmAnaArticolo();
                    f._strArtCod = sArt;
                    f.ShowDialog();
                    if (f._tabArt != null && f._tabArt.Rows.Count > 0 && (string)f._tabArt.Rows[0]["art_cod"] != sArt)
                    {
                        x["leg_leg"] = (string)f._tabArt.Rows[0]["art_cod"];
                        x["art_des"] = (string)f._tabArt.Rows[0]["art_des"];
                        x["art_tra"] = (string)f._tabArt.Rows[0]["art_tra"];
                    }
                    else if (f._tabArt != null && f._tabArt.Rows.Count > 0)
                        x["art_tra"] = (string)f._tabArt.Rows[0]["art_tra"];
                    x["LegMdy"] = "S";
                }
            }

        }

        private void dgv1_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            dgv1.Rows[e.RowIndex].Cells["Modificato"].Value = "S";
            Totali();
        }

        private void dgv1_EditingControlShowing(object sender, DataGridViewEditingControlShowingEventArgs e)
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

        private void dgv1_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            if (dgv1.Columns[e.ColumnIndex].Name == "Lotto")
            {
                if (!DBNull.Value.Equals(dgv1.Rows[e.RowIndex].Cells[e.ColumnIndex].Value))
                {
                    string s = (string)dgv1.Rows[e.RowIndex].Cells[e.ColumnIndex].Value;
                    if (s != "S")
                        e.CellStyle.BackColor = Color.Red;
                }
            }
        }

        private void txt_KeyPress(object sender, KeyPressEventArgs e)
        {
            if ((sender as TextBox).Text != null && (sender as TextBox).Text.Trim() != "")
            {
                string s = dgv1.CurrentCell.OwningColumn.Name.ToString();

                if (s == "Inc. %")
                {
                    if (e.KeyChar == '.')
                        e.KeyChar = ',';
                }
            }
        }

        private void Totali()
        {
            decimal d = 0;

            foreach(DataRow y in ((DataTable)dgv1.DataSource).Rows)
            {
                if (DBNull.Value.Equals(y["leg_ann"]) || !(Boolean)y["leg_ann"])
                {
                    if (_clsFun.Numerico(y["leg_inc"]))
                        d += (decimal)y["leg_inc"];
                }
            }

            lblInc.Text = d.ToString("#0.00");
        }

        private Boolean Salva()
        {
            Boolean b = true;

            if (dgv1.CurrentCell is DataGridViewCell)
            {
                dgv1.CommitEdit(DataGridViewDataErrorContexts.Commit);
            }

            string s = "";
            DataRow x;
            DataRow[] j;

            s = "SELECT * FROM AnaArtLegami WHERE leg_tip='" + _strTip + "' AND leg_art='" + _strArt + "'";
            DataTable tLeg = _clsFun.FillTabSql("AnaArtLegami", s, false, _strConSql);

            DataTable t = (DataTable)dgv1.DataSource;

            ArrayList aWhe = new ArrayList();
            //aWhe.Add("leg_tip");
            //aWhe.Add("leg_art");
            aWhe.Add("leg_idx");
            ArrayList aExl = new ArrayList();

            foreach (DataRow y in t.Rows)
            {
                if ((string)y["LegMdy"] != "")
                {
                    x = tLeg.NewRow();
                    x["leg_art"] = y["leg_art"];
                    x["leg_tip"] = y["leg_tip"];
                    x["leg_leg"] = y["leg_leg"];
                    x["leg_ann"] = y["leg_ann"];

                    if (DBNull.Value.Equals(y["leg_idx"]) || (int)y["leg_idx"] == 0)
                        s = _clsFun.SqlInsertRow("AnaArtLegami", tLeg, y);
                    else
                    {
                        //s = "leg_idx=" + Convert.ToInt32(y["leg_idx"]);
                        //j = tLeg.Select(s);
                        //s = "";
                        //if (j.Length > 0)
                        //    s = _clsFun.SqlUpdRowIdx("AnaArtLegami", tLeg, j[0], y, aExl);

                        s = "leg_idx=" + Convert.ToInt32(y["leg_idx"]);
                        j = tLeg.Select(s);
                        s = "";
                        if (j.Length > 0)
                            s = _clsFun.SqlUpdRowIdx("AnaArtLegami", tLeg, j[0], y, aExl);
                    }

                    if (s != "")
                        _clsFun.SqlWrite(s, _strConSql);
                }
            }

            return b;
        }

        private void btnOK_Click(object sender, EventArgs e)
        {
            Boolean b = true;

            if (dgv1.DataSource == null || ((DataTable)dgv1.DataSource).Rows.Count == 0)
            {
                MessageBox.Show("Non ci sono articoli da generare barcode!", "GENERAZIONE BARCODE", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                b = false;
            }
            if (_strKey.Trim().Length > 3)
            {
                MessageBox.Show("Lunghezza codice lotto maggiore di 3!", "GENERAZIONE BARCODE", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                b = false;
            }

            if (b)
            {
                string s = _clsFun.ParGet(clsDefine.enuParametri.Par031Lotti2Pos, _strConSql);
                string sPrf = "";
                string sEanPes = "";
                string[] a = s.Split(',');
                if (a.Length > 1 && a[1] != "")
                {
                    sPrf = a[1];
                    if (a.Length > 4 && a[4] != "")
                        sEanPes = a[4];
                }
                else
                {
                    MessageBox.Show("Prefisso Lotto/barcode non definito in configurazione!", "GENERAZIONE BARCODE", MessageBoxButtons.OK, MessageBoxIcon.Question);
                    b = false;
                }

                if(b)
                {
                    if (MessageBox.Show("Confermi la generazione dei barcode per articolo?", "GENERAZIONE BARCODE", MessageBoxButtons.YesNo, MessageBoxIcon.Question) != DialogResult.Yes)
                        b = false;
                }

                if (b)
                    GenBarcode(sPrf, sEanPes);
            }
        }

        private void GenBarcode(string strPrf, string strEanPes)
        {
            string s = "";

            DataTable t = (DataTable)dgv1.DataSource;

            foreach(DataRow y in t.Rows)
            {
                if (!(Boolean)y["leg_ann"])
                {
                    string sEan = strPrf + _strKey + y["leg_rig"] + new string('0', 6);

                    s = "SELECT * FROM AnaBarcode WHERE ean_ean='" + sEan + "'";
                    DataTable tTmp = _clsFun.FillTabSql("AnaBarcode", s, true, _strConSql);
                    if (tTmp.Rows.Count > 0)
                    {
                        DataRow x = tTmp.NewRow();

                        foreach (DataColumn c in tTmp.Columns)
                            x[c.ColumnName] = tTmp.Rows[0][c.ColumnName];

                        x["ean_art"] = y["leg_leg"];
                        x["ean_tip"] = "L";             //Lotto
                        x["ean_dti"] = DateTime.Today;
                        x["ean_dtm"] = DateTime.Today;
                        x["ean_bil"] = true;
                        x["ean_ann"] = false;
                        if (strEanPes == "EANPESO")
                            x["ean_ecp"] = true;

                        s = _clsFun.SqlUpdRowIdx("AnaBarcode", tTmp, tTmp.Rows[0], x, null);
                    }
                    else
                    {
                        DataRow x = tTmp.NewRow();
                        x["ean_art"] = y["leg_leg"];
                        x["ean_ean"] = sEan;
                        x["ean_ann"] = false;
                        x["ean_qta"] = 0;
                        x["ean_prv"] = 0;
                        x["ean_dti"] = DateTime.Today;
                        x["ean_dtm"] = DateTime.Today;
                        x["ean_bil"] = true;
                        x["ean_ecp"] = false;

                        if (strEanPes == "EANPESO")
                            x["ean_ecp"] = true;

                        x["ean_pun"] = 0;
                        x["ean_tip"] = "L";

                        s = _clsFun.SqlInsertRow("AnaBarcode", tTmp, x);
                    }

                    _clsFun.FileLog("LOTTO EAN", "", s);

                    if (s != "")
                    {
                        _clsFun.SqlWrite(s, _strConSql);

                    }

                    //s = "UPDATE AnaBarcode SET ean_bil=0 WHERE ";
                    //s += "ean_art='" + y["leg_leg"] +"' AND ";
                    //s += "ean_ean<>'" + sEan +"'";
                    //_clsFun.SqlWrite(s, _strConSql);

                    s = "UPDATE AnaBarcode SET ean_bil=0 WHERE ";
                    s += "ean_art='" + y["leg_leg"] + "' AND ";
                    s += "ean_ean<>'" + sEan + "'";
                    _clsFun.SqlWrite(s, _strConSql);

                    _clsVar.Variazioni((string)y["leg_leg"], "Forza", "Lotti articolo", _clsDef.VARPOS);
                }
            }
        }

        private void cancellaAnnullatiToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("Confermi la cancellazione delle righe annullate?", "CANCELLAZIONE ELIMINATI", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
                string s = "DELETE FROM AnaArtLegami WHERE leg_ann=1 AND leg_tip='" + _strTip + "' AND leg_art='" + _strArt + "'";
                _clsFun.SqlWrite(s, _strConSql);
                FillDati();
            }
        }


    }
}
