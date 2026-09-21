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
    public partial class frmSeekArt : Form
    {
        private const string SEEKARTCOD = "ARTCOD";
        private const string SEEKARTDES = "ARTDES";
        private const string SEEKEANCOD = "EANCOD";
        private const string SEEKARFCOD = "ARFCOD";
        private const string SEEKARTPLU = "ARTPLU";
        private const string SEEKARTTAS = "ARTTAS";
        private const string SEEKREPBIL = "REPBIL";
        private const string SEEKREPART = "REPART";
        private const string SEEKARTFOR = "ARTFOR";
        private const string SEEKARTETI = "ARTETI";
        private const string SEEKARTWEB = "ARTWEB";
        private const string SEEKARTCEL = "ARTCEL";
        private const string SEEKARTCAN = "ARTCAN";
        private const string SEEKARTLAST = "SEEKLAST";

        private const string TABANAART = "AnaArticoli";
        private const string TABLISACQ = "GesLisAcquisto";
        
        clsDefine _clsDef = new clsDefine();
        clsFuncs _clsFun = new clsFuncs();
        clsQuery _clsQry = new clsQuery();
        clsVariazioni _clsVar = new clsVariazioni();

        private string _strConSql = "";
        public string _strSeek = "";

        //private static string _strUltTip = "";
        //public string _strUltSeek = "";

        public Boolean _bolAnaArt = false;
        public Boolean _bolValori = false;

        public DataTable _tabArt = new DataTable();

        public frmSeekArt()
        {
            InitializeComponent();
            new clsGesGraph().SetGraph(this, 0);
            _strConSql = _clsFun.ConSql("");
            FillTabs();
            SetDgv1();
        }

        private void frmSeekArt_Load(object sender, EventArgs e)
        {
            ApplyModernStyle();
            txtArtCod.Text = "";
            txtEanCod.Text = "";
            txtArtDes.Text = "";
            txtArfCod.Text = "";
            txtArtPlu.Text = "";
            txtArtTas.Text = "";

            if(_strSeek != "")
            {
                if(_clsFun.Numerico(_strSeek))
                {
                    if (_strSeek.Length < 8)
                    {
                        txtArtCod.Text = _strSeek;
                        txtArtCod.Select();
                    }
                    else
                    {
                        txtEanCod.Text = _strSeek;
                        txtEanCod.Select();
                    }
                }
                else
                {
                    txtArtDes.Text = _strSeek;
                    if (FillDati(SEEKARTDES))
                        dgv1.Select();
                    else
                        txtArtDes.Select();
                }
            }

            txtArtDes.Select();
        }

        private void frmSeekArt_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Escape)
            {
                e.Handled = true;
                e.SuppressKeyPress = true;
                Esci();
            }
        }

        private void esciToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Esci();
        }

        private void Esci()
        {
            this.Close();
            //this.Dispose();
        }

        private void ApplyModernStyle()
        {
            this.Font = new Font("Segoe UI", 9F, FontStyle.Regular, GraphicsUnit.Point);
            this.BackColor = Color.FromArgb(248, 250, 252); // Soft light slate background

            panel1.BackColor = Color.FromArgb(241, 245, 249); // Card panel
            panel1.BorderStyle = BorderStyle.FixedSingle;

            if (lblTrovati != null)
            {
                lblTrovati.Font = new Font("Segoe UI", 9.5F, FontStyle.Bold);
                lblTrovati.ForeColor = Color.FromArgb(30, 41, 59);
            }
        }

        private void SetDgv1()
        {
            dgv1.AutoGenerateColumns = false;
            dgv1.AllowUserToAddRows = false;
            dgv1.ReadOnly = false;
            dgv1.AllowUserToDeleteRows = false;

            // Abilita DoubleBuffering per eliminare il flickering
            try
            {
                typeof(DataGridView).InvokeMember("DoubleBuffered",
                    System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.SetProperty,
                    null, dgv1, new object[] { true });
            }
            catch { }

            // Stile Header moderno (Slate Dark Blue)
            dgv1.EnableHeadersVisualStyles = false;
            dgv1.BorderStyle = BorderStyle.None;
            dgv1.CellBorderStyle = DataGridViewCellBorderStyle.SingleHorizontal;
            dgv1.GridColor = Color.FromArgb(226, 232, 240);

            dgv1.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(30, 41, 59); // #1E293B
            dgv1.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
            dgv1.ColumnHeadersDefaultCellStyle.Font = new Font("Segoe UI", 9.5F, FontStyle.Bold);
            dgv1.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft;
            dgv1.ColumnHeadersHeight = 32;
            dgv1.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing;

            // Stile Righe
            dgv1.DefaultCellStyle.Font = new Font("Segoe UI", 9F, FontStyle.Regular);
            dgv1.DefaultCellStyle.BackColor = Color.White;
            dgv1.DefaultCellStyle.ForeColor = Color.FromArgb(15, 23, 42);
            dgv1.DefaultCellStyle.SelectionBackColor = Color.FromArgb(14, 116, 144); // Sleek Cyan/Teal #0E7490
            dgv1.DefaultCellStyle.SelectionForeColor = Color.White;
            dgv1.RowTemplate.Height = 28;

            dgv1.AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(248, 250, 252); // Soft alternating background
            dgv1.AlternatingRowsDefaultCellStyle.SelectionBackColor = Color.FromArgb(14, 116, 144);
            dgv1.AlternatingRowsDefaultCellStyle.SelectionForeColor = Color.White;

            DataGridViewTextBoxColumn cTbc;
            DataGridViewButtonColumn cBtn;

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "art_cod";
            cTbc.Name = "Codice";
            cTbc.Width = 85;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = true;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "art_des";
            cTbc.Name = "Descrizione";
            cTbc.AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
            cTbc.MinimumWidth = 220;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = true;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "tmp_cos";
            cTbc.Name = "Costo";
            cTbc.Width = 75;
            cTbc.ValueType = typeof(string);
            cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            cTbc.DefaultCellStyle.Format = "###,##0.000";
            cTbc.ReadOnly = true;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "tmp_prv";
            cTbc.Name = "P.vendita";
            cTbc.Width = 75;
            cTbc.ValueType = typeof(string);
            cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            cTbc.DefaultCellStyle.Format = "###,##0.00";
            cTbc.ReadOnly = true;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "art_sta";
            cTbc.Name = "Stato";
            cTbc.Width = 45;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = true;
            cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dgv1.Columns.Add(cTbc);

            cBtn = new DataGridViewButtonColumn();
            cBtn.UseColumnTextForButtonValue = true;
            cBtn.FlatStyle = FlatStyle.Flat;
            cBtn.Text = "Var";
            cBtn.Width = 38;
            cBtn.ValueType = typeof(string);
            cBtn.ReadOnly = false;
            cBtn.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            cBtn.ToolTipText = "Gen variazione POS";
            dgv1.Columns.Add(cBtn);

            cBtn = new DataGridViewButtonColumn();
            cBtn.UseColumnTextForButtonValue = true;
            cBtn.FlatStyle = FlatStyle.Flat;
            cBtn.Text = "Etv";
            cBtn.Width = 38;
            cBtn.ValueType = typeof(string);
            cBtn.ReadOnly = false;
            cBtn.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            cBtn.ToolTipText = "Gen variazione etichette";
            dgv1.Columns.Add(cBtn);
        }

        private void txtArtCod_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Return)
                FillDati(SEEKARTCOD);
        }

        private void txtArtDes_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Return)
                FillDati(SEEKARTDES);
        }

        private void txtEanCod_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Return)
                FillDati(SEEKEANCOD);
        }

        private void txtArfCod_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Return)
                FillDati(SEEKARFCOD);
        }

        private void txtArtPlu_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Return)
                FillDati(SEEKARTPLU);
        }

        private void txtArtTas_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Return)
                FillDati(SEEKARTTAS);
        }
               
        private void cmbArtRep_SelectionChangeCommitted(object sender, EventArgs e)
        {
            FillDati(SEEKREPART);
        }

        private void cmbArtReb_SelectionChangeCommitted(object sender, EventArgs e)
        {
            FillDati(SEEKREPBIL);
        }

        private void cmbArtFor_SelectionChangeCommitted(object sender, EventArgs e)
        {
            FillDati(SEEKARTFOR);
        }

        private void cmbArtEti_SelectionChangeCommitted(object sender, EventArgs e)
        {
            FillDati(SEEKARTETI);
        }

        private void rdbArtWeb_CheckedChanged(object sender, EventArgs e)
        {
            FillDati(SEEKARTWEB);
        }
        private void rdbArtCel_CheckedChanged(object sender, EventArgs e)
        {
            FillDati(SEEKARTCEL);
        }
        private void rdbArtCan_CheckedChanged(object sender, EventArgs e)
        {
            FillDati(SEEKARTCAN);
        }
        private void nuovoArticoloToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //string s = _clsFun.ParGet(clsDefine.enuParametri.Par036ArtInsVeloce, _strConSql);
            //if (s.Length > 0 && s.Substring(0, 1) == "S")
            //{
            //    frmAnaArtNuovo f = new frmAnaArtNuovo();
            //    f._strEan = txtEanCod.Text;
            //    f.ShowDialog();
            //}
            //else
            //{
            frmAnaArticolo f = new frmAnaArticolo();
            f.ShowDialog();
            //}
        }

        private void FillTabs()
        {
            string p = "";
            string s = "";
            DataRow x;
            DataTable t;

            p = "TabRepBilance";
            s = "SELECT * FROM " + p + " WHERE tab_ann=0 ORDER BY tab_des";
            t = _clsFun.FillTabSql(p, s, false, _strConSql);
            x = t.NewRow();
            x["tab_cod"] = "";
            x["tab_des"] = "  Non definito";
            t.Rows.InsertAt(x, 0);
            cmbArtReb.DataSource = t;
            cmbArtReb.DisplayMember = "tab_des";
            cmbArtReb.ValueMember = "tab_cod";
            cmbArtReb.SelectedValue = "";

            p = "TabReparti";
            s = "SELECT * FROM TabReparti ORDER BY tab_des";
            t = _clsFun.FillTabSql(p, s, false, _strConSql);
            x = t.NewRow();
            x["tab_cod"] = "";
            x["tab_des"] = "  Non definito";
            t.Rows.InsertAt(x, 0);
            cmbArtRep.DataSource = t;
            cmbArtRep.DisplayMember = "tab_des";
            cmbArtRep.ValueMember = "tab_cod";
            cmbArtRep.SelectedValue = "";

            p = "TabEtichette";
            s = "SELECT * FROM TabEtichette WHERE tab_ann=0 ORDER BY tab_des";
            t = _clsFun.FillTabSql(p, s, false, _strConSql);
            x = t.NewRow();
            x["tab_cod"] = "";
            x["tab_des"] = "  Non definito";
            t.Rows.InsertAt(x, 0);
            cmbArtEti.DataSource = t;
            cmbArtEti.DisplayMember = "tab_des";
            cmbArtEti.ValueMember = "tab_cod";
            cmbArtEti.SelectedValue = "";

            p = "AnaFornitori";
            s = "SELECT * FROM AnaFornitori ORDER BY for_des";
            t = _clsFun.FillTabSql(p, s, false, _strConSql);
            x = t.NewRow();
            x["for_cod"] = "";
            x["for_des"] = "  Non definito";
            t.Rows.InsertAt(x, 0);
            cmbArtFor.DataSource = t;
            cmbArtFor.DisplayMember = "for_des";
            cmbArtFor.ValueMember = "for_cod";
            cmbArtFor.SelectedValue = "";
        }

        private Boolean FillDati(string strTip)
        {
            string s = ""; 
            string sWhe = "";
            string sSql = "";
            Boolean b = true;
            //Boolean bLast = false;

            if (strTip == SEEKARTLAST)
            {
                //bLast = true;

                if (strTip == SEEKARTCOD)
                    txtArtCod.Text = s;
                if (strTip == SEEKARTDES)
                    txtArtDes.Text = s;

            }
            if (strTip == SEEKARTCOD)
            {
                if (txtArtCod.Text == "")
                    b = false;
                else
                {
                    txtArtCod.Text = txtArtCod.Text.PadLeft(7, Convert.ToChar("0"));
                    sWhe = "art_cod='" + txtArtCod.Text + "'";
                }
            }
            else if (strTip == SEEKARTDES)
            {
                if (txtArtDes.Text == "")
                    b = false;
                else
                {
                    if(txtArtDes.Text.Contains("+"))
                    {
                        s = txtArtDes.Text;

                        string[] a = s.Split('+');
                        s = "";
                        for(int i = 0; i < a.Length; i++)
                        {
                            s += "art_des LIKE '%" + a[i] + "%' AND ";
                        }
                        sWhe = s.Substring(0, s.Length - 4);
                    }
                    else
                        sWhe = "art_des LIKE '%" + txtArtDes.Text.Trim() + "%'";
                }
            }
            else if (strTip == SEEKREPART)
            {
                if (cmbArtRep.SelectedValue.ToString() == "")
                    b = false;
                else
                    sWhe = "art_rep = '" + cmbArtRep.SelectedValue.ToString() + "'";
            }
            else if (strTip == SEEKARTWEB)
            {
                if (rdbArtWeb.Checked)
                    sWhe = "art_web = 1 ";
            }
            else if (strTip == SEEKARTCEL)
            {
                if (rdbArtCel.Checked)
                    sWhe = "art_cel = 1 ";
            }
            else if (strTip == SEEKARTCAN)
            {
                if (rdbArtCan.Checked)
                    sWhe = "art_sta = 'C' ";
            }
            else if (strTip == SEEKREPBIL)
            {
                if (cmbArtReb.SelectedValue.ToString() == "")
                    b = false;
                else
                    sWhe = "art_reb = '" + cmbArtReb.SelectedValue.ToString() + "'";
            }
            else if (strTip == SEEKARTPLU)
            {
                string sVal = txtArtPlu.Text.Trim();
                if (string.IsNullOrEmpty(sVal))
                    b = false;
                else
                {
                    if (int.TryParse(sVal, out int iPlu))
                    {
                        sWhe = "(art_plu = '" + sVal.Replace("'", "''") + "' OR art_plu = '" + iPlu.ToString() + "' OR art_plu = '" + iPlu.ToString("D4") + "' OR art_plu = '" + iPlu.ToString("D6") + "' OR (ISNUMERIC(art_plu) = 1 AND CAST(art_plu AS int) = " + iPlu + "))";
                    }
                    else
                    {
                        sWhe = "art_plu = '" + sVal.Replace("'", "''") + "'";
                    }
                }
            }
            else if (strTip == SEEKARTTAS)
            {
                string sVal = txtArtTas.Text.Trim();
                if (string.IsNullOrEmpty(sVal))
                    b = false;
                else
                {
                    if (int.TryParse(sVal, out int iTas))
                    {
                        sWhe = "(art_tas = '" + sVal.Replace("'", "''") + "' OR art_tas = '" + iTas.ToString() + "' OR art_tas = '" + iTas.ToString("D2") + "' OR art_tas = '" + iTas.ToString("D3") + "' OR art_tas = '" + iTas.ToString("D4") + "' OR (ISNUMERIC(art_tas) = 1 AND CAST(art_tas AS int) = " + iTas + "))";
                    }
                    else
                    {
                        sWhe = "art_tas = '" + sVal.Replace("'", "''") + "'";
                    }
                }
            }
            else if (strTip == SEEKARTETI)
            {
                if (cmbArtEti.SelectedValue.ToString() == "")
                    b = false;
                else
                    sWhe = "art_eti = '" + cmbArtEti.SelectedValue.ToString() + "'";
            }
            else if (strTip == SEEKEANCOD)
            {
                if (txtEanCod.Text == "")
                    b = false;
                else
                {
                    sSql = "SELECT ean_art, ean_ean, AnaArticoli.* ";
                    sSql += "FROM AnaBarcode ";
                    sSql += "LEFT OUTER JOIN AnaArticoli ON AnaBarcode.ean_art = AnaArticoli.art_cod ";
                    sSql += "WHERE ean_ean='" + txtEanCod.Text + "'";

                    DataTable t = _clsFun.FillTabSql(TABANAART, sSql, false, _strConSql);

                    if (t.Rows.Count == 0)
                        b = false;
                    else
                        sWhe = "art_cod='" + t.Rows[0]["ean_art"] + "'";
                }
            }
            else if (strTip == SEEKARFCOD)
            {
                if (txtArfCod.Text == "")
                    b = false;
                else
                {
                    sSql = "SELECT lia_art, lia_ann FROM GesLisAcquisto WHERE lia_arf = '" + txtArfCod.Text + "' ORDER BY lia_art, lia_dti DESC";
                    DataTable t = _clsFun.FillTabSql(TABLISACQ, sSql, false, _strConSql);

                    if (t.Rows.Count == 0)
                        b = false;
                    else
                    {
                        //_strUltSeek = strTip + ";" + txtArfCod.Text;

                        s = "";
                        foreach (DataRow y in t.Rows)
                        {
                            if (s != (string)y["lia_art"])
                            {
                                if (DBNull.Value.Equals(y["lia_ann"]) || !(Boolean)y["lia_ann"])
                                {
                                    sWhe += "art_cod='" + y["lia_art"] + "' OR ";
                                    s = (string)y["lia_art"];
                                }
                            }
                        }
                        if (sWhe != "")
                            sWhe = sWhe.Substring(0, sWhe.Length - 4);
                    }
                }
            }
            else if (strTip == SEEKARTFOR)
            {
                if (cmbArtFor.SelectedValue == null || cmbArtFor.SelectedValue.ToString() == "")
                    b = false;
                else
                {
                    sSql = "SELECT lia_art, lia_ann FROM GesLisAcquisto WHERE lia_for = '" + cmbArtFor.SelectedValue.ToString() + "' ORDER BY lia_art, lia_dti DESC";
                    DataTable t = _clsFun.FillTabSql(TABLISACQ, sSql, false, _strConSql);

                    if (t.Rows.Count == 0)
                        b = false;
                    else
                    {
                        //_strUltSeek = strTip + ";" + txtArfCod.Text;

                        s = "";
                        foreach (DataRow y in t.Rows)
                        {
                            if (s != (string)y["lia_art"])
                            {
                                if (DBNull.Value.Equals(y["lia_ann"]) || !(Boolean)y["lia_ann"])
                                {
                                    sWhe += "art_cod='" + y["lia_art"] + "' OR ";
                                    s = (string)y["lia_art"];
                                }
                            }
                        }
                        if (sWhe != "")
                            sWhe = sWhe.Substring(0, sWhe.Length - 4);
                    }
                }
            }

            if(b && sWhe != "")
            {
                string todaySql = _clsFun.DaySql(DateTime.Today);
                string sqlCos = @"
                    OUTER APPLY (
                        SELECT TOP 1 lia_cos 
                        FROM GesLisAcquisto 
                        WHERE lia_art = AnaArticoli.art_cod 
                          AND (lia_ann=0 OR lia_ann IS NULL) 
                          AND (lia_dti <= " + todaySql + @" OR lia_dti IS NULL)
                        ORDER BY lia_dti DESC
                    ) c ";

                string sqlPrv = @"
                    OUTER APPLY (
                        SELECT TOP 1 liv_prv 
                        FROM GesLisVendita 
                        WHERE liv_art = AnaArticoli.art_cod 
                          AND liv_lis = '" + _clsDef.LISPOS + @"' 
                          AND (liv_ann=0 OR liv_ann IS NULL) 
                          AND (liv_dti <= " + todaySql + @" OR liv_dti IS NULL)
                        ORDER BY liv_dti DESC
                    ) v ";

                sSql = "SELECT art_cod, art_des, art_sta, art_eqp, ISNULL(c.lia_cos, 0) AS tmp_cos, ISNULL(v.liv_prv, 0) AS tmp_prv FROM AnaArticoli " + sqlCos + sqlPrv + " WHERE " + sWhe;

                DataTable t = _clsFun.FillTabSql(TABANAART, sSql, false, _strConSql);

                dgv1.DataSource = t;
                dgv1.Select();

                lblTrovati.Text = t.Rows.Count.ToString();
                _bolValori = true;
            }

            return dgv1.RowCount > 0;
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
            if (dgv1.DataSource != null)
            {
                CurrencyManager cm = dgv1.BindingContext[dgv1.DataSource, this.dgv1.DataMember] as CurrencyManager;
                if (cm.Position >= 0)
                {
                    DataRowView r = cm.List[dgv1.CurrentRow.Index] as DataRowView;
                    DataRow x = r.Row;

                    _tabArt = new clsGenTabTmp().TabTmpArt("OrdArt");
                    DataRow k = _tabArt.NewRow();
                    k["tmp_art"] = x["art_cod"];
                    k["tmp_ard"] = x["art_des"];
                    k["tmp_eqp"] = x["art_eqp"];
                    //k["tmp_sta"] = x["art_sta"];
                    //k["tmp_rep"] = x["art_rep"];
                    //k["tmp_iva"] = x["art_iva"];
                    //k["tmp_umi"] = x["art_umi"];
                    //k["tmp_bil"] = x["art_bil"];
                    _tabArt.Rows.Add(k);

                    if (_bolAnaArt)
                    {
                        frmAnaArticolo f = new frmAnaArticolo();
                        f._tabArt = _tabArt.Copy();
                        f.ShowDialog();

                        //if (f._bolCodNew && f._strArtCod != _clsDef.CODNEW)
                        //    ArtSeek("DES");
                        //x["art_des"] = f._strArtDes;
                    }
                    else
                        Esci();
                }
                else
                    Esci();
            }
        }

        private void dgv1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.ColumnIndex == 5)
            {
                CurrencyManager cm = dgv1.BindingContext[dgv1.DataSource, dgv1.DataMember] as CurrencyManager;
                if (cm.Position >= 0)
                {
                    DataRowView r = cm.List[dgv1.CurrentRow.Index] as DataRowView;
                    DataRow x = r.Row;

                    if ((string)x["art_sta"] != _clsDef.STAATT && (string)x["art_sta"] != _clsDef.STACAN && (string)x["art_sta"] != _clsDef.STAREP)
                        MessageBox.Show("Articolo non attivo", "CONTROLLO ARTICOLO", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    else
                    {
                        _clsVar.Variazioni((string)x["art_cod"], "Forza", "Ricerca articolo", "POS");

                        string sPar016PathDivNegozi = _clsFun.ParGet(clsDefine.enuParametri.Par016PathDivNegozi, _strConSql);
                        if (sPar016PathDivNegozi != "")
                        {
                            string sFor = "";

                            if (cmbArtFor.SelectedValue != null && cmbArtFor.SelectedValue.ToString() != "")
                                sFor = cmbArtFor.SelectedValue.ToString();

                            ArrayList a = new ArrayList();
                            a.Add((string)x["art_cod"]);
                            _clsVar.DivNegArticoli(sPar016PathDivNegozi, a, null, "", sFor);
                        }

                        MessageBox.Show("Variazione POS generata con successo per l'articolo " + (string)x["art_cod"] + " - " + (string)x["art_des"] + ".", "VARIAZIONE POS GENERATA", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
            }
            if (e.ColumnIndex == 6)
            {
                CurrencyManager cm = dgv1.BindingContext[dgv1.DataSource, dgv1.DataMember] as CurrencyManager;
                if (cm.Position >= 0)
                {
                    DataRowView r = cm.List[dgv1.CurrentRow.Index] as DataRowView;
                    DataRow x = r.Row;

                    if ((string)x["art_sta"] != _clsDef.STAATT && (string)x["art_sta"] != _clsDef.STAREP)
                        MessageBox.Show("Articolo non attivo", "CONTROLLO ARTICOLO", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    else
                    {
                        _clsVar.Variazioni((string)x["art_cod"], "Forza", "Ricerca articolo", "ETI");
                        MessageBox.Show("Variazione Etichette generata con successo per l'articolo " + (string)x["art_cod"] + " - " + (string)x["art_des"] + ".", "VARIAZIONE ETICHETTE GENERATA", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
            }
        }

        private void txtArtPlu_TextChanged(object sender, EventArgs e)
        {

        }

     }
}
