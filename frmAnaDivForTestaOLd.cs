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
    public partial class frmAnaDivForTestaOLd : Form
    {
        clsDefine _clsDef = new clsDefine();
        clsFuncs _clsFun = new clsFuncs();

        private string _strConSql = "";

        public DataRow _rowTrk;
        public DataRow _rowTrt;

        //public string _strCod = "";
        //public string _strTip = "";
        //public string _strTid = "";
        //public string _strTis = "";
        //public string _strRes = "";

        public frmAnaDivForTestaOLd()
        {
            InitializeComponent();
            new clsGesGraph().SetGraph(this, 0);
        }

        private void frmAnaDivForTesta_Load(object sender, EventArgs e)
        {
            _strConSql = _clsFun.ConSql("");

            FillDati();
        }

        private void esciToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Esci();
        }
        private void frmAnaDivForTesta_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Escape)
                Esci();
        }
        private void Esci()
        {
            //Salva();
            //this.Close();

            Boolean b = Salva();
            if (!b)
            {
                if (MessageBox.Show("Abbandono delle modifiche, confermi?", this.Text, MessageBoxButtons.YesNo, MessageBoxIcon.Question)
                      == DialogResult.Yes)
                    this.Close();
            }
            else
                this.Close();


        }

        private void FillDati()
        {
            //string s = "SELECT * FROM AnaForDivTestate WHERE trt_for='" + _strTip + "' AND trt_cod='" + _strCod + "'";
            //DataTable t = _clsFun.FillTabSql("AnaForDivTestate", s, true, _strConSql);

            //if (t.Rows.Count > 0)
            //{
            //    lblTrkCod.Text = _strTip;
            //    txtTrkDes.Text = _strTid;
            //    lblTrtCod.Text = _strCod;
            //    txtTrtFil.Text = (string)t.Rows[0]["trt_fil"];
            //    txtTrtPth.Text = (string)t.Rows[0]["trt_pth"];
            //    txtTrtTip.Text = (string)t.Rows[0]["trt_tip"];
            //    txtTrtLeg.Text = (string)t.Rows[0]["trt_leg"];
            //}

            if (_rowTrk != null)
            {
                if (!DBNull.Value.Equals(_rowTrk["trk_cod"]))
                    lblTrkCod.Text = (string)_rowTrk["trk_cod"];        // _strTip;
                if (!DBNull.Value.Equals(_rowTrk["trk_des"]))
                    txtTrkDes.Text = (string)_rowTrk["trk_des"];        //_strTid;
                if (!DBNull.Value.Equals(_rowTrk["trk_sta"]))
                    txtTrkSta.Text = (string)_rowTrk["trk_sta"];        //_strTis;
            }

            if (lblTrkCod.Text == "")
            {
                lblTrkCod.Text = _clsDef.CODNEW;
                txtTrkSta.Text = _clsDef.STAATT;
                txtTrkDes.Text = "";

            }

            if (_rowTrt != null)
            {
                if (!DBNull.Value.Equals(_rowTrt["trt_cod"]))
                    lblTrtCod.Text = (string)_rowTrt["trt_cod"];
                if (!DBNull.Value.Equals(_rowTrt["trt_fil"]))
                    txtTrtFil.Text = (string)_rowTrt["trt_fil"];
                if (!DBNull.Value.Equals(_rowTrt["trt_pth"]))
                    txtTrtPth.Text = (string)_rowTrt["trt_pth"];
                if (!DBNull.Value.Equals(_rowTrt["trt_tip"]))
                    txtTrtTip.Text = (string)_rowTrt["trt_tip"];
                if (!DBNull.Value.Equals(_rowTrt["trt_leg"]))
                    txtTrtLeg.Text = (string)_rowTrt["trt_leg"];
                if (!DBNull.Value.Equals(_rowTrt["trt_reg"]))
                    txtTrtReg.Text = (string)_rowTrt["trt_reg"];
            }
            else
                lblTrtCod.Text = _clsDef.CODNEW;
        }

        private Boolean Salva()
        {
            string s = "";
            string sMsg = "";

            if (lblTrkCod.Text == "")
                sMsg += "Codice tracciato non definito" + _clsDef.CRLF;
            if (txtTrkDes.Text == "")
                sMsg += "Descrizione tracciato non definito" + _clsDef.CRLF;
            if (lblTrtCod.Text == "")
                sMsg += "Codice file tracciato non definito" + _clsDef.CRLF;
            if (txtTrtFil.Text == "")
                sMsg += "File non definito" + _clsDef.CRLF;
            if (txtTrtPth.Text == "")
                sMsg += "Path file non definito" + _clsDef.CRLF;
            if (txtTrtTip.Text == "")
                sMsg += "Tipo file non definito" + _clsDef.CRLF;

            if (sMsg != "")
                MessageBox.Show(sMsg,"CONTROLLO TRACCIATO", MessageBoxButtons.OK, MessageBoxIcon.Error);
            else
            {
                string sCod = lblTrkCod.Text;

                if (sCod == _clsDef.CODNEW)
                {
                    sCod = TipoNuovo();
                    lblTrkCod.Text = sCod;
                }

                s = "SELECT * FROM AnaForDivTipi WHERE trk_cod ='" + sCod + "'";
                DataTable t = _clsFun.FillTabSql("AnaForDivTipi", s, true, _strConSql);
                DataRow x = t.NewRow();
                x["trk_cod"] = sCod;
                x["trk_des"] = txtTrkDes.Text;
                x["trk_sta"] = txtTrkSta.Text;

                if (t.Rows.Count > 0)
                {
                    x["trk_idx"] = t.Rows[0]["trk_idx"];
                    s = _clsFun.SqlUpdRowIdx("AnaForDivTipi", t, t.Rows[0], x, null);
                }
                else
                    s = _clsFun.SqlInsertRow("AnaForDivTipi", t, x);
                if (s != "")
                    _clsFun.SqlWrite(s, _strConSql);

                _rowTrk = x;

                //_rowTrk["trk_cod"] = sCod;
                //_rowTrk["trk_des"] = txtTrkDes.Text;
                //_rowTrk["trk_sta"] = txtTrkSta.Text;

                sCod = lblTrtCod.Text;

                if (sCod == _clsDef.CODNEW)
                {
                    sCod = "0001";

                    s = "SELECT * FROM AnaForDivTestate WHERE trt_for='" + lblTrkCod.Text + "' ORDER BY trt_cod DESC";
                    DataTable tTrt = _clsFun.FillTabSql("AnaForDivTestate", s, true, _strConSql);
                    if (tTrt.Rows.Count > 0)
                    {
                        s = (string)tTrt.Rows[0]["trt_cod"];
                        sCod = (Convert.ToInt16(s) + 1).ToString("0000");
                    }
                }

                s = "SELECT * FROM AnaForDivTestate WHERE trt_for='" + lblTrkCod.Text + "' AND trt_cod='" + sCod + "'";
                t = _clsFun.FillTabSql("AnaForDivTestate", s, true, _strConSql);
                x = t.NewRow();
                x["trt_for"] = lblTrkCod.Text;
                x["trt_cod"] = sCod;
                x["trt_fil"] = txtTrtFil.Text;
                x["trt_pth"] = txtTrtPth.Text;
                x["trt_tip"] = txtTrtTip.Text;
                x["trt_leg"] = txtTrtLeg.Text;
                x["trt_reg"] = txtTrtReg.Text;
                x["trt_ann"] = false;

                if (t.Rows.Count > 0)
                {
                    x["trt_idx"] = t.Rows[0]["trt_idx"];
                    s = _clsFun.SqlUpdRowIdx("AnaForDivTestate", t, t.Rows[0], x, null);
                }
                else
                    s = _clsFun.SqlInsertRow("AnaForDivTestate", t, x);
                if (s != "")
                    _clsFun.SqlWrite(s, _strConSql);

                //_strRes = lblTrkCod.Text + ";" + lblTrtCod.Text;

                _rowTrt = x;

                //_rowTrt["trt_for"] = lblTrkCod.Text;
                //_rowTrt["trt_cod"] = sCod;
                //_rowTrt["trt_fil"] = txtTrtFil.Text;
                //_rowTrt["trt_pth"] = txtTrtPth.Text;
                //_rowTrt["trt_tip"] = txtTrtTip.Text;
                //_rowTrt["trt_leg"] = txtTrtLeg.Text;
                //_rowTrt["trt_ann"] = false;
            }

            return sMsg == "";
        }

        private void btnNew_Click(object sender, EventArgs e)
        {
            lblTrkCod.Text = _clsDef.CODNEW;
            txtTrkDes.Text = "";
            txtTrkSta.Text = _clsDef.STAATT;
        }

        private string TipoNuovo()
        {
            string sCod = "00001";

            string s = "SELECT trk_cod FROM AnaForDivTipi ORDER BY trk_cod DESC";

            DataTable t = _clsFun.FillTabSql("AnaForDivTipi", s, false, _strConSql);

            if(t.Rows.Count > 0)
                sCod = (Convert.ToInt32( t.Rows[0]["trk_cod"])+1).ToString("00000");

            return sCod;
        }

    }
}
