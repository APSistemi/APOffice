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
    public partial class frmUtyCtrlJournal : Form
    {
        clsDefine _clsDef = new clsDefine();
        clsFuncs _clsFun = new clsFuncs();
        clsQuery _clsQry = new clsQuery();

        private string _strConSql = "";
        private string _strConSqlSta = "";

        public frmUtyCtrlJournal()
        {
            InitializeComponent();

            _strConSql = _clsFun.ConSql("");
            _strConSqlSta = _clsFun.ConSql("3");
        }

        private void frmUtyCtrlJournal_Load(object sender, EventArgs e)
        {
            dtpDay.Value = new DateTime(2019, 11, 16);
            //dtpDay.Value = DateTime.Today;
        }

        private void esciToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void btnJou_Click(object sender, EventArgs e)
        {
            OpenFileDialog o = new OpenFileDialog();
            DialogResult dr = o.ShowDialog();

            if (dr == System.Windows.Forms.DialogResult.OK)
            {
                try
                {
                    lblPath.Text = o.FileName;
                }
                catch (Exception ex)
                {
                    MessageBox.Show("File non disponibile: " + o.FileName + " - " + ex.Message);
                }
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            FillJournal();
            FillDati();
        }

        private void xFillJournal() //Funziona con Porta PO
        {
            string sDtp = dtpDay.Value.ToString("dd/MM/yy");

            DataTable tTmp = new clsGenTabTmp().TabTmpJounal("TmpArt");
            DataColumn[] keys = new DataColumn[2];
            keys[0] = tTmp.Columns["tmp_art"];

            string sFil = lblPath.Text;
            string s = "";
            string sRig = "";
            decimal dTot = 0;

            using (StreamReader sr = new StreamReader(sFil))
            {
                decimal dImp = 0;
                int i = 0;

                while ((sRig = sr.ReadLine()) != null)
                {
                    string sDay = "";
                    string sOra = "";
                    string sSco = "";

                    if (sRig.Length > 0 && sRig.Substring(0, 6) == "TOTALE")
                    {
                        dImp = Convert.ToDecimal(sRig.Substring(22, 14));
                    }

                    if (sRig.Length > 0 && sRig.Substring(2, 1) == "/")
                    {
                        sDay = sRig.Substring(0, 8);
                        sOra = sRig.Substring(9, 5);
                        sSco = sRig.Substring(28, 4);

                        if (sDay == sDtp)
                        {
                            sSco = (Convert.ToInt16(sSco) - 3).ToString("0000");

                            if (sDay != "" && dImp > 0)
                            {
                                i++;

                                DataRow x = tTmp.NewRow();
                                x["tmp_day"] = sDay;
                                x["tmp_ora"] = sOra;
                                x["tmp_sco"] = sSco;
                                x["tmp_imj"] = dImp;
                                x["tmp_num"] = i.ToString("0000");
                                tTmp.Rows.Add(x);

                                dTot += dImp;

                                dImp = 0;
                            }
                        }
                    }

                }
                sr.Close();
                sr.Dispose();

                dgv1.DataSource = tTmp;

                lblToj.Text = dTot.ToString("#,##0.00");

            }
            // MessageBox.Show("Fatto!");
        }

        private void yFillJournal()
        {
            string sDtp = dtpDay.Value.ToString("dd/MM/yy");

            DataTable tTmp = new clsGenTabTmp().TabTmpJounal("TmpArt");
            DataColumn[] keys = new DataColumn[2];
            keys[0] = tTmp.Columns["tmp_sco"];

            string sFil = lblPath.Text;
            string s = "";
            string sRig = "";
            decimal dTot = 0;

            using (StreamReader sr = new StreamReader(sFil))
            {
                DataRow[] j;
                decimal dImp = 0;
                int i = 0;

                decimal dI_A = 0;
                decimal dI_B = 0;
                decimal dI_C = 0;
                decimal dI_D = 0;

                DataRow x = tTmp.NewRow();
                x["tmp_day"] = "Totali";
                x["tmp_ora"] = "";
                x["tmp_sco"] = "XXXX";
                x["tmp_imj"] = 0;
                x["tmp_i_A"] = 0;
                x["tmp_i_B"] = 0;
                x["tmp_i_C"] = 0;
                x["tmp_i_D"] = 0;
                x["tmp_num"] = 0;
                tTmp.Rows.Add(x);

                while ((sRig = sr.ReadLine()) != null)
                {
                    string sDay = "";
                    string sOra = "";
                    string sSco = "";

                    if (sRig.Length > 39 && sRig.Substring(20, 6) == "HEADER")
                    {
                        dImp = 0;
                        dI_A = 0;
                        dI_B = 0;
                        dI_C = 0;
                        dI_D = 0;
                    }
                    else
                    {
                        if (sRig.Length > 6 && sRig.Substring(0, 6) == "TOTALE")
                        {
                            dImp = Convert.ToDecimal(sRig.Substring(22, 22));
                        }
                        else
                        {
                            if (sRig.Length > 43 && sRig.Substring(41, 3) != "IVA")
                            {
                                if (sRig.Length > 43 && sRig.Substring(43, 1) == "A")
                                    dI_A += Convert.ToDecimal(sRig.Substring(33, 9));
                                else if (sRig.Length > 43 && sRig.Substring(43, 1) == "B")
                                    dI_B += Convert.ToDecimal(sRig.Substring(33, 9));
                                else if (sRig.Length > 43 && sRig.Substring(43, 1) == "C")
                                    dI_C += Convert.ToDecimal(sRig.Substring(33, 9));
                                else if (sRig.Length > 43 && sRig.Substring(43, 1) == "D")
                                    dI_D += Convert.ToDecimal(sRig.Substring(33, 9));
                            }
                        }

                        if (sRig.Length > 31 && sRig.Substring(31, 3) == "DOC")
                        {
                            sDay = sRig.Substring(0, 8);
                            sOra = sRig.Substring(9, 5);
                            sSco = sRig.Substring(40, 4);

                            if (sDay == sDtp)
                            {
                                if (sDay != "" && dImp > 0)
                                {
                                    i++;

                                    x = tTmp.NewRow();
                                    x["tmp_day"] = sDay;
                                    x["tmp_ora"] = sOra;
                                    x["tmp_sco"] = sSco;
                                    x["tmp_imj"] = dImp;
                                    x["tmp_i_A"] = dI_A;
                                    x["tmp_i_B"] = dI_B;
                                    x["tmp_i_C"] = dI_C;
                                    x["tmp_i_D"] = dI_D;
                                    x["tmp_num"] = i.ToString("0000");
                                    tTmp.Rows.Add(x);

                                    dTot += dImp;

                                    j = tTmp.Select("tmp_sco='XXXX'");
                                    j[0]["tmp_imj"] = (decimal)j[0]["tmp_imj"] + dImp;
                                    j[0]["tmp_i_A"] = (decimal)j[0]["tmp_i_A"] + dI_A;
                                    j[0]["tmp_i_B"] = (decimal)j[0]["tmp_i_B"] + dI_B;
                                    j[0]["tmp_i_C"] = (decimal)j[0]["tmp_i_C"] + dI_C;
                                    j[0]["tmp_i_D"] = (decimal)j[0]["tmp_i_D"] + dI_D;

                                }

                                dImp = 0;
                                dI_A = 0;
                                dI_B = 0;
                                dI_C = 0;
                                dI_D = 0;

                            }
                        }
                    }
                }
                sr.Close();
                sr.Dispose();

                dgv1.DataSource = tTmp;

                lblToj.Text = dTot.ToString("#,##0.00");

            }
           // MessageBox.Show("Fatto!");
        }

        private void FillJournal()
        {
            string sDtp = dtpDay.Value.ToString("dd/MM/yyyy");

            DataTable tTmp = new clsGenTabTmp().TabTmpJounal("TmpArt");
            DataColumn[] keys = new DataColumn[2];
            keys[0] = tTmp.Columns["tmp_sco"];

            string sFil = lblPath.Text;
            string s = "";
            string sRig = "";
            decimal dTot = 0;

            using (StreamReader sr = new StreamReader(sFil))
            {
                DataRow[] j;
                decimal dImp = 0;
                int i = 0;

                decimal dI_A = 0;
                decimal dI_B = 0;
                decimal dI_C = 0;
                decimal dI_D = 0;

                DataRow x = tTmp.NewRow();
                x["tmp_day"] = "Totali";
                x["tmp_ora"] = "";
                x["tmp_sco"] = "XXXX";
                x["tmp_imj"] = 0;
                x["tmp_i_A"] = 0;
                x["tmp_i_B"] = 0;
                x["tmp_i_C"] = 0;
                x["tmp_i_D"] = 0;
                x["tmp_num"] = 0;
                tTmp.Rows.Add(x);

                string sDay = "";
                string sOra = "";
                string sSco = "";

                while ((sRig = sr.ReadLine()) != null)
                {

                    //if (sRig.Length > 39 && sRig.Substring(20, 6) == "HEADER")
                    if (sRig.Length > 39 && sRig.Substring(12, 22) == "DOCUMENTO  COMMERCIALE")
                    {
                        dImp = 0;
                        dI_A = 0;
                        dI_B = 0;
                        dI_C = 0;
                        dI_D = 0;
                    }
                    else
                    {
                        if (sRig.Length > 30 && sRig.Substring(0, 18) == "TOTALE COMPLESSIVO")
                        {
                            dImp = Convert.ToDecimal(sRig.Substring(22, 24));
                        }
                        else if (sRig.Length > 31 && sRig.Substring(21, 4) == "2019")
                        {
                            sDay = sRig.Substring(15, 10).Replace("-", "/");
                            sOra = sRig.Substring(26, 5);
                        }
                        else if (sRig.Length > 43 && sRig.Substring(0, 10) == "di cui IVA")
                            Console.WriteLine("Salto");
                        else if (sRig.Length > 43 && sRig.Substring(0, 9) == "Pagamento")
                            Console.WriteLine("Salto");
                        else if (sRig.Length > 43 && sRig.Substring(0, 12) == "Non riscosso")
                            Console.WriteLine("Salto");
                        else if (sRig.Length > 43 && sRig.Substring(0, 5) == "Resto")
                            Console.WriteLine("Salto");
                        else if (sRig.Length > 43 && sRig.Substring(0, 7) == "Importo")
                            Console.WriteLine("Salto");
                        else if (sRig.Length > 43 && sRig.Substring(0, 11) == "DESCRIZIONE")
                            Console.WriteLine("Salto");
                        else if (sRig.Length > 43 && sRig.Substring(26, 2) == "22")
                            dI_A += Convert.ToDecimal(sRig.Substring(37, 9));
                        else if (sRig.Length > 43 && sRig.Substring(26, 2) == "10")
                            dI_B += Convert.ToDecimal(sRig.Substring(37, 9));
                        else if (sRig.Length > 43 && sRig.Substring(26, 2) == " 4")
                            dI_C += Convert.ToDecimal(sRig.Substring(37, 9));
                        else if (sRig.Length > 43 && sRig.Substring(26, 2) == " 5")
                            dI_D += Convert.ToDecimal(sRig.Substring(37, 9));
                        else if (sRig.Length > 31 && sRig.Substring(12, 11) == "DOCUMENTO N")
                        {
                            sSco = sRig.Substring(25, 4);

                            if (sDay == sDtp)
                            {
                                if (sDay != "" && dImp > 0)
                                {
                                    i++;

                                    x = tTmp.NewRow();
                                    x["tmp_day"] = sDay;
                                    x["tmp_ora"] = sOra;
                                    x["tmp_sco"] = sSco;
                                    x["tmp_imj"] = dImp;
                                    x["tmp_i_A"] = dI_A;
                                    x["tmp_i_B"] = dI_B;
                                    x["tmp_i_C"] = dI_C;
                                    x["tmp_i_D"] = dI_D;
                                    x["tmp_num"] = i.ToString("0000");
                                    tTmp.Rows.Add(x);

                                    dTot += dImp;

                                    j = tTmp.Select("tmp_sco='XXXX'");
                                    j[0]["tmp_imj"] = (decimal)j[0]["tmp_imj"] + dImp;
                                    j[0]["tmp_i_A"] = (decimal)j[0]["tmp_i_A"] + dI_A;
                                    j[0]["tmp_i_B"] = (decimal)j[0]["tmp_i_B"] + dI_B;
                                    j[0]["tmp_i_C"] = (decimal)j[0]["tmp_i_C"] + dI_C;
                                    j[0]["tmp_i_D"] = (decimal)j[0]["tmp_i_D"] + dI_D;

                                }

                                dImp = 0;
                                dI_A = 0;
                                dI_B = 0;
                                dI_C = 0;
                                dI_D = 0;
                            }
                        }
                    }
                }
                sr.Close();
                sr.Dispose();

                dgv1.DataSource = tTmp;

                lblToj.Text = dTot.ToString("#,##0.00");

            }
            // MessageBox.Show("Fatto!");
        }

        private void FillDati()
        {
            DataRow[] j;
            DataRow[] jj;
            int i = 0;

            string s = "SELECT * FROM GesNegVet WHERE vet_cau='SCO' AND vet_day=" + _clsFun.DaySql(dtpDay.Value) + " AND vet_neg='001' AND vet_pos='" + txtPos.Text + "'";

            DataTable tVet = _clsFun.FillTabSql("GesNegVet", s, false, _strConSqlSta);
            DataColumn[] keys = new DataColumn[1];
            keys[0] = tVet.Columns["vet_sco"];
            tVet.PrimaryKey = keys;

            s = "SELECT * FROM GesNegVen WHERE ven_cau='SCO' AND ven_day=" + _clsFun.DaySql(dtpDay.Value) + " AND ven_neg='001' AND ven_pos='" + txtPos.Text + "'";

            DataTable tVen = _clsFun.FillTabSql("GesNegVen", s, false, _strConSqlSta);
            keys = new DataColumn[6];
            keys[0] = tVen.Columns["ven_ora"];
            keys[1] = tVen.Columns["ven_pos"];
            keys[2] = tVen.Columns["ven_sco"];
            keys[3] = tVen.Columns["ven_sct"];
            keys[4] = tVen.Columns["ven_ean"];
            keys[5] = tVen.Columns["ven_art"];
            tVen.PrimaryKey = keys;

            DataTable t = (DataTable)dgv1.DataSource;

            decimal dTot = 0;

            foreach(DataRow y in t.Rows)
            {
                if ((string)y["tmp_sco"] != "XXXX")
                { 
                    i++;

                    s = i.ToString("00000");

                    j = tVet.Select("vet_sco='" + s + "'");
                    if (j.Length > 0)
                    {

                        y["tmp_ims"] = (decimal)j[0]["vet_imp"];
                        y["tmp_dif"] = (decimal)y["tmp_ims"] - (decimal)y["tmp_imj"];

                        dTot += (decimal)j[0]["vet_imp"];

                        string sSco = (string)j[0]["vet_sco"];

                        j = tVen.Select("ven_sco='" + sSco + "'");

                        for (int ii = 0; ii < j.Length; ii++)
                        {
                            if ((string)j[ii]["ven_iva"] == "022")
                            {
                                y["tmp_s_A"] = (decimal)y["tmp_s_A"] + (decimal)j[ii]["ven_ven"];
                                jj = t.Select("tmp_sco='XXXX'");
                                jj[0]["tmp_s_A"] = (decimal)jj[0]["tmp_s_A"] + (decimal)j[ii]["ven_ven"];

                            }
                            else if ((string)j[ii]["ven_iva"] == "010")
                            {
                                y["tmp_s_B"] = (decimal)y["tmp_s_B"] + (decimal)j[ii]["ven_ven"];
                                jj = t.Select("tmp_sco='XXXX'");
                                jj[0]["tmp_s_B"] = (decimal)jj[0]["tmp_s_B"] + (decimal)j[ii]["ven_ven"];
                            }
                            else if ((string)j[ii]["ven_iva"] == "004")
                            {
                                y["tmp_s_C"] = (decimal)y["tmp_s_C"] + (decimal)j[ii]["ven_ven"];
                                jj = t.Select("tmp_sco='XXXX'");
                                jj[0]["tmp_s_C"] = (decimal)jj[0]["tmp_s_C"] + (decimal)j[ii]["ven_ven"];
                            }
                            else if ((string)j[ii]["ven_iva"] == "005")
                            {
                                y["tmp_s_D"] = (decimal)y["tmp_s_D"] + (decimal)j[ii]["ven_ven"];
                                jj = t.Select("tmp_sco='XXXX'");
                                jj[0]["tmp_s_D"] = (decimal)jj[0]["tmp_s_D"] + (decimal)j[ii]["ven_ven"];
                            }

                            //y["tmp_i_A"] = (decimal)y["tmp_i_A"] + dI_A;
                            //y["tmp_i_B"] = (decimal)y["tmp_i_B"] + dI_B;
                            //y["tmp_i_C"] = (decimal)y["tmp_i_C"] + dI_C;
                            //y["tmp_i_D"] = (decimal)y["tmp_i_D"] + dI_D;
                        }
                    }
                }
            }

            lblTom.Text = dTot.ToString("#,##0.00");


        }
    }
}
