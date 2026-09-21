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
    public partial class frmGesDocLottiDettaglio : Form
    {
        clsDefine _clsDef = new clsDefine();
        clsFuncs _clsFun = new clsFuncs();
        clsQuery _clsQry = new clsQuery();

        clsResize _form_resize;
        private string _strFrmPar = "";

        private string _strConSql = "";

        public string _strTip = "";
        public string _strArt = "";
        public string _strArd = "";
        public DataRow _rowLot;

        public string _strMovFat= "";           //Tipo documento
        public string _strMovYea= "";           //Anno
        public string _strMftMov = "";          //Numero movimento / fattura
        public string _strMftNum = "";          //Numero documento
        public string _strMovRow = "";          //Numero riga

        private DataTable _tabNaz = new DataTable("TabNaz");
        private DataTable _tabPrv = new DataTable("TabPrv");

        public frmGesDocLottiDettaglio()
        {
            InitializeComponent();

            _form_resize = new clsResize(this);
            this.Load += _Load;
            this.Resize += _Resize;

            new clsGesGraph().SetGraph(this, 0);
        }
        private void _Load(object sender, EventArgs e)
        {
            _form_resize._get_initial_size();
        }
        private void _Resize(object sender, EventArgs e)
        {
            _form_resize._resize();
        }
        private void frmGesDocLottiDettaglio_Load(object sender, EventArgs e)
        {
            Video();

            if (_strTip == "RowDoc")
                pnlDoc.Enabled = false;
            else
                selezioneLottoEsistenteToolStripMenuItem.Enabled = false;

            FillTab();
            FillDati();

            _strConSql = _clsFun.ConSql("");
        }
        private void Video()
        {
            _form_resize._get_initial_size();
            _strFrmPar = _clsQry.ParForm(this, "R", "");

            this.CenterToScreen();
        }
        private void esciToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Esci();
        }
        private void frmGesDocLottiDettaglio_KeyDown(object sender, KeyEventArgs e)
        {
            if(e.KeyCode == Keys.Escape)
                Esci();
        }
        private void Esci()
        {
            string sMsg = Salva();
            if(sMsg != "")
            {
                if (MessageBox.Show( sMsg + _clsDef.CRLF + "Abbandono delle modifiche, confermi?", "CONTROLLO LOTTO", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                    this.Close();
            }
            else
                this.Close();
        }

        private void FillTab()
        {
            string p = "AnaFornitori";
            string s = "";

            s = "SELECT for_cod, for_des FROM AnaFornitori ORDER BY for_des";
            DataTable t = _clsFun.FillTabSql(p, s, false, _strConSql);
            DataRow x = t.NewRow();
            x["for_cod"] = "";
            x["for_des"] = "  Non definito";
            t.Rows.InsertAt(x, 0);
            cmbFor.DataSource = t;
            cmbFor.DisplayMember = "for_des";
            cmbFor.ValueMember = "for_cod";
            cmbFor.SelectedValue = "";

            p = "TabLottoTipo";
            s = "SELECT * FROM TabLottoTipo WHERE tab_ann=0 ORDER BY tab_des";
            t = _clsFun.FillTabSql(p, s, false, _strConSql);
            x = t.NewRow();
            x["tab_cod"] = "";
            x["tab_des"] = "  Non definito";
            t.Rows.InsertAt(x, 0);
            cmbTip.DataSource = t;
            cmbTip.DisplayMember = "tab_des";
            cmbTip.ValueMember = "tab_cod";
            cmbTip.SelectedValue = "";

            p = "TabStato";
            s = "SELECT * FROM TabStato WHERE tab_lot=1 ORDER BY tab_cod";
            t = _clsFun.FillTabSql(p, s, false, _strConSql);
            x = t.NewRow();
            x["tab_cod"] = "";
            x["tab_des"] = "  Non definito";
            t.Rows.InsertAt(x, 0);
            cmbSta.DataSource = t;
            cmbSta.DisplayMember = "tab_des";
            cmbSta.ValueMember = "tab_cod";
            cmbSta.SelectedValue = "";

            p = "TabOrigine";
            s = "SELECT * FROM TabOrigine ORDER BY tab_des";
            _tabNaz = _clsFun.FillTabSql(p, s, false, _strConSql);
            //x = _tabNaz.NewRow();
            //x["tab_cod"] = "";
            //x["tab_des"] = "  Non definito";
            //_tabNaz.Rows.InsertAt(x, 0);

            p = "TabProvince";
            s = "SELECT * FROM TabProvince ORDER BY tab_des";
            _tabPrv = _clsFun.FillTabSql(p, s, false, _strConSql);
            //x = _tabNaz.NewRow();
            //x["tab_cod"] = "";
            //x["tab_des"] = "  Non definito";
            //_tabNaz.Rows.InsertAt(x, 0);
        }

        private void FillDati()
        {
            if (_clsFun.Numerico(_rowLot["lot_cod"], "0123456789") && Convert.ToInt16(_rowLot["lot_cod"]) > 0)
                txtCod.Text = ((string)_rowLot["lot_cod"]).PadLeft(5, Convert.ToChar('0'));
            else
                txtCod.Text = "";
            cmbFor.SelectedValue = (string)_rowLot["lot_for"];
            txtNdo.Text = (string)_rowLot["lot_ndo"];
            dayDdo.Value = (DateTime)_rowLot["lot_ddo"];
            cmbTip.SelectedValue = (string)_rowLot["lot_tip"];

            txtDes.Text = (string)_rowLot["lot_des"];
            if (txtDes.Text == "")
                txtDes.Text = _strArd;

            txtNat.Text = (string)_rowLot["lot_nat"];
            txtAllNaz.Text = (string)_rowLot["lot_all"];
            
            txtAllPrv.Text = "";
            if (!DBNull.Value.Equals(_rowLot["lot_alp"]))
                txtAllPrv.Text = (string)_rowLot["lot_alp"];

            txtMac.Text = (string)_rowLot["lot_mac"];
            txtSez.Text = (string)_rowLot["lot_sez"];
            txtLotLot.Text = (string)_rowLot["lot_lot"];
            txtMacNaz.Text = (string)_rowLot["lot_mna"];
            txtSezBol.Text = (string)_rowLot["lot_sbo"];
            cmbSta.SelectedValue = (string)_rowLot["lot_sta"];

            if (!DBNull.Value.Equals(_rowLot["lot_mdt"]))
                txtMacDta.Text = (string)_rowLot["lot_mdt"];

            if(DBNull.Value.Equals(_rowLot["lot_art"]))
                _rowLot["lot_art"] = "";
            if (_strArt == "" && (string)_rowLot["lot_art"] != "")
                _strArt = (string)_rowLot["lot_art"];

            _strMovFat  = (string)_rowLot["lot_dot"];
            _strMovYea = (string)_rowLot["lot_doy"];
            _strMftNum = (string)_rowLot["lot_don"];
            
            //_strMftMov = "";
            //if (!DBNull.Value.Equals(_rowLot["lot_dom"]))
            _strMftMov = (string)_rowLot["lot_dom"];

            _strMovRow = (string)_rowLot["lot_dor"];
        }

        private string Salva()
        {
            string s = "";
            string sMsg = "";

            if (txtCod.Text == "")
                sMsg = "Codice lotto non definito";
            else
            {
                txtCod.Text = txtCod.Text.PadLeft(5, Convert.ToChar('0'));

                _rowLot["lot_yea"] = dayDdo.Value.Year.ToString();
                _rowLot["lot_cod"] = txtCod.Text;
                _rowLot["lot_for"] = cmbFor.SelectedValue;
                _rowLot["ForDes"] = cmbFor.Text;
                _rowLot["lot_ndo"] = txtNdo.Text;
                _rowLot["lot_ddo"] = dayDdo.Value;
                _rowLot["lot_tip"] = cmbTip.SelectedValue;
                _rowLot["lot_des"] = txtDes.Text;
                _rowLot["lot_nat"] = txtNat.Text;
                _rowLot["lot_all"] = txtAllNaz.Text;
                _rowLot["lot_alp"] = txtAllPrv.Text;
                _rowLot["lot_mac"] = txtMac.Text;
                _rowLot["lot_sez"] = txtSez.Text;
                _rowLot["lot_sbo"] = txtSezBol.Text;

                _rowLot["lot_lot"] = txtLotLot.Text;
                _rowLot["lot_mna"] = txtMacNaz.Text;
                _rowLot["lot_mdt"] = txtMacDta.Text;

                _rowLot["lot_dot"] = _strMovFat;
                _rowLot["lot_doy"] = _strMovYea;
                _rowLot["lot_don"] = _strMftNum;
                _rowLot["lot_dor"] = _strMovRow;
                _rowLot["lot_art"] = _strArt;

                _rowLot["lot_dom"] = _strMftMov;


                //if (_strMovFat == "F")
                //    xx["lot_dor"] = (string)x["mov_rfa"];
                //else
                //    xx["lot_dor"] = (string)x["mov_rmo"];
                //xx["lot_dor"] = _strMovRow;
                //xx["lot_art"] = _strArt;

                _rowLot["lot_sta"] = "";
                if (cmbSta.SelectedValue != null)
                    _rowLot["lot_sta"] = cmbSta.SelectedValue.ToString();

                if ((string)_rowLot["LotMdy"] == "")
                    _rowLot["LotMdy"] = "S";

                s = "SELECT * FROM GesDocLotti WHERE lot_yea = '" + dayDdo.Value.Year.ToString() + "' AND lot_cod='" + txtCod.Text + "'";
                DataTable tt = _clsFun.FillTabSql("GesDocLotti", s, true, _strConSql);

                /* Salva Lotto*/

                DataRow xx = _rowLot;

                if(tt.Rows.Count > 0)
                    xx["lot_idx"] = tt.Rows[0]["lot_idx"];

                xx["lot_dot"] = _rowLot["lot_dot"];
                xx["lot_doy"] = _rowLot["lot_doy"];
                xx["lot_don"] = _rowLot["lot_don"];
                xx["lot_dom"] = _rowLot["lot_dom"];

                //if (_strMovFat == "F")
                //    xx["lot_dor"] = (string)x["mov_rfa"];
                //else
                //    xx["lot_dor"] = (string)x["mov_rmo"];

                //xx["lot_dor"] = _strMovRow;
                //xx["lot_art"] = _strArt;

                xx["lot_dor"] = _rowLot["lot_dor"];
                xx["lot_art"] = _rowLot["lot_art"];

                if (DBNull.Value.Equals(xx["lot_idx"]) || (int)xx["lot_idx"] == 0)
                    s = _clsFun.SqlInsertRow("GesDocLotti", tt, xx);
                else
                {
                    ArrayList aWhe = new ArrayList();
                    aWhe.Add("lot_idx");
                    //ArrayList aExl = new ArrayList();
                    s = "lot_idx=" + Convert.ToInt32(xx["lot_idx"]);
                    DataRow[] j = tt.Select(s);
                    s = "";
                    if (j.Length > 0)
                        s = _clsFun.SqlUpdRowIdx("GesDocLotti", tt, j[0], xx, null);
                }

                if (s != "")
                    _clsFun.SqlWrite(s, _strConSql);

                s = "SELECT * FROM GesDocLotti WHERE lot_yea = '" + dayDdo.Value.Year.ToString() + "' AND lot_cod='" + txtCod.Text + "'";
                tt = _clsFun.FillTabSql("GesDocLotti", s, false, _strConSql);
                if (tt.Rows.Count > 1)
                    MessageBox.Show("Codice lotto presente su più righe!", "CONTROLLO LOTTO", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }

            return sMsg;
        }

        private void txt_DoubleClick(object sender, EventArgs e)
        {
            TextBox txt = (TextBox)sender;

            frmUtyTastiera f = new frmUtyTastiera();
            f._strStr = txt.Text;
            f.ShowDialog();
            txt.Text = f._strStr;

            if (txt.Name == "txtNat" || txt.Name == "txtAllNaz" || txt.Name == "txtMacNaz" || txt.Name == "txtSez")
                txt.Text = NazSeek(txt.Text);
            else if (txt.Name == "txtAllPrv")
                txt.Text = PrvSeek(txt.Text);
        }

        private void txtNaz_KeyDown(object sender, KeyEventArgs e)
        {
            if(e.KeyCode == Keys.Return)
            {
                TextBox txt = (TextBox)sender;
                if(txt.Text != "")
                {
                    if (txt.Name == "txtAllPrv")
                        txt.Text = PrvSeek(txt.Text);
                    else
                        txt.Text = NazSeek(txt.Text);
                }
            }
        }

        private string NazSeek(string strTxt)
        {
            string s = "tab_des LIKE'%" + strTxt + "%'";

            DataRow[] j = _tabNaz.Select(s);

            s = strTxt;
            if (j.Length > 0)
            {
                if (j.Length == 1)
                    s = (string)j[0]["tab_des"];
                else
                {
                    frmUtyTab f = new frmUtyTab();
                    f._strTip = "LotNaz";
                    f._tabTab = _tabNaz;
                    f._strSql = "tab_des LIKE'%" + strTxt + "%'";
                    f.ShowDialog();
                    //if (f._strRes != "")
                    s = f._strRes;
                }
            }
            return s;
        }

        private string PrvSeek(string strTxt)
        {
            string s = "tab_des LIKE'%" + strTxt + "%'";

            DataRow[] j = _tabPrv.Select(s);

            s = strTxt;
            if (j.Length > 0)
            {
                if (j.Length == 1)
                    s = (string)j[0]["tab_cod"];
                else
                {
                    frmUtyTab f = new frmUtyTab();
                    f._strTip = "LotNaz";
                    f._tabTab = _tabPrv;
                    f._strSql = "tab_des LIKE'%" + strTxt + "%'";
                    f.ShowDialog();
                    //if (f._strRes != "")
                    s = f._strCod;
                }
            }
            return s;
        }

        private void selezioneLottoEsistenteToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Salva();
            frmSeekDocLotti f = new frmSeekDocLotti();
            f._strTip = _strTip;
            f._strSelect = "S";
            f.ShowDialog();

            if (f._strRes != "")
            {
                if (txtLotLot.Text == "" || MessageBox.Show("Sostituzione lotto, confermi?", "CONTROLLO LOTTO", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                {
                    string sYea = f._strYea;
                    string sLot = f._strRes.Trim();

                    string s = "SELECT * FROM GesDocLotti WHERE lot_yea='" + sYea + "' AND lot_cod='" + sLot + "'";
                    DataTable t = _clsFun.FillTabSql("GesDocLotti", s, true, _strConSql);
                    if (t.Rows.Count > 0)
                    {
                        foreach (DataColumn c in t.Columns)
                            _rowLot[c.ColumnName] = t.Rows[0][c.ColumnName];

                        _rowLot["lot_doy"] = sYea;
                        FillDati();
                    }
                }
            }
        }

        private void btnNew_Click(object sender, EventArgs e)
        {
            if(txtCod.Text != "")
                MessageBox.Show("Codice lotto già definito!", "CONTROLLO LOTTO", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            else
                txtCod.Text = _clsFun.NewNum(DateTime.Today.Year.ToString("0000"), clsDefine.enuNumeratori.NumGesDocLotti, 5, _strConSql);
        }

        private void btnTag_Click(object sender, EventArgs e)
        {
            if (txtCod.Text == "" || !_clsFun.Numerico(txtCod.Text) || Convert.ToInt16(txtCod.Text) == 0)
                MessageBox.Show("Codice lotto non definito!", "CONTROLLO LOTTO", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            else if(_strArt == "")
                MessageBox.Show("Codice articolo non definito!", "CONTROLLO LOTTO", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            else
            {
                Salva();
                frmAnaArtLegami f = new frmAnaArtLegami();
                f._strTip = "LOT";
                f._strArt = _strArt;
                f._strArd = _strArd;
                f._strKey = txtCod.Text.Substring(2);
                f.ShowDialog();
            }
        }

        private void btnCtrl_Click(object sender, EventArgs e)
        {            
            if (txtCod.Text == "")
                MessageBox.Show("Codice lotto non definito!", "CONTROLLO LOTTO", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            else if (_strArt == "")
                MessageBox.Show("Codice articolo non definito!", "CONTROLLO LOTTO", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            else
            {
                Salva();

                frmGesStatMovLegami f = new frmGesStatMovLegami();
                f._strTip = "LOT";
                f._strArt = _strArt;
                f._strArd = _strArd;
                f._rowLot = _rowLot;
                f.ShowDialog();
            }
        }

        private void txtMacDta_KeyDown(object sender, KeyEventArgs e)
        {
            frmUtyDay f = new frmUtyDay();
            f.ShowDialog();
            txtMacDta.Text = f._strDay;
        }

        private void txtMacDta_MouseClick(object sender, MouseEventArgs e)
        {
            frmUtyDay f = new frmUtyDay();
            f.ShowDialog();
            txtMacDta.Text = f._strDay;
        }

        private void txtCod_Validated(object sender, EventArgs e)
        {
            if(txtCod.Text != "")
            {
                if (!_clsFun.Numerico(txtCod.Text, "0123456789"))
                    txtCod.Text = "";
            }
        }

        private void btnPrn_Click(object sender, EventArgs e)
        {
            if(!_rowLot.Table.Columns.Contains("MovQkg"))
                MessageBox.Show("Stampabile solo dal dettaglio movimento di carico", "CONTROLLO LOTTO", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            else if (txtCod.Text == "")
                MessageBox.Show("Codice lotto non definito!", "CONTROLLO LOTTO", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            else if (cmbTip.SelectedValue == null || cmbTip.SelectedValue.ToString() == "")
                MessageBox.Show("Tipo lotto non definito!", "CONTROLLO LOTTO", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            else
            {
                Salva();

                string sTip = "02";
                string sFor = cmbFor.SelectedValue.ToString();
                string sFod = cmbFor.Text;
                string sInd = "";
                string sPar = "";
                int iScaGio = 0;
                decimal dPes = (decimal)_rowLot["MovQkg"];

                DataTable t = _clsQry.ArtSeek(_strArt, "SEEK");
                if (t.Rows.Count > 0 && !DBNull.Value.Equals(t.Rows[0]["tmp_gsc"]) && _clsFun.Numerico((string)t.Rows[0]["tmp_gsc"]))
                    iScaGio = Convert.ToInt16(t.Rows[0]["tmp_gsc"]);

                string sLot = cmbTip.SelectedValue.ToString().PadRight(2, Convert.ToChar(' ')) + "-" + txtCod.Text + "-" + _strMovYea;

                frmGesMovIngLabel04 f = new frmGesMovIngLabel04();
                f._strEtiTipo = sTip;
                f._strConSql = _strConSql;
                f._strCliCod = sFor;
                f._strCliDes = sFod;
                f._strCliInd = "";
                f._strCliPar = "";
                f._strArtCod = _strArt;
                f._strArtDes = _strArd;
                f._strArtUmi = "KG";
                f._decMovQkg = dPes;
                f._decMovPrv = 0;
                f._decMovImp = 0;
                f._dayArtSca = DateTime.Today.AddDays(iScaGio);
                f._strLotCod = sLot; // txtCod.Text;         // txtLot.Text;
                f.ShowDialog();
            }
        }

        private void btnMov_Click(object sender, EventArgs e)
        {
            if (txtCod.Text == "")
                MessageBox.Show("Codice lotto non definito!", "CONTROLLO LOTTO", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            else if (_strArt == "")
                MessageBox.Show("Codice articolo non definito!", "CONTROLLO LOTTO", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            else
            {
                Salva();

                frmGesStatMovLotto f = new frmGesStatMovLotto();
                f._strTip = "LOT";
                f._strArt = _strArt;
                f._strArd = _strArd;
                f._rowLot = _rowLot;
                f.ShowDialog();
            }

        }

    }
}
