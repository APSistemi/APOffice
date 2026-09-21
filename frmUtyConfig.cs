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
    public partial class frmUtyConfig : Form
    {
        clsDefine _clsDef = new clsDefine();
        clsFuncs _clsFun = new clsFuncs();

        private string _strConSql = "";
        private string _strConSqlSta = "";
        private string _strConSqlStaSto = "";
        private string _strConSqlLog = "";

        private const string TABANA = "AnaConfigurazione";

        public frmUtyConfig()
        {
            InitializeComponent();
            new clsGesGraph().SetGraph(this, 0);
        }

        private void frmUtiConfig_Load(object sender, EventArgs e)
        {
            _strConSql = _clsFun.ConSql("");
            _strConSqlSta = _clsFun.ConSql("3");
            _strConSqlLog = _clsFun.ConSql("4");
            _strConSqlStaSto = _clsFun.ConSql("5");

            string s = _clsFun.FileIni("R", clsDefine.enuIni.Ini09CodiceAzienda, "");
            if (s == "")
                s = "001";
            txtCnfCod.Text = s;

            FillTabs();
            FillDati();
        }

        private void frmUtiConfig_KeyDown(object sender, KeyEventArgs e)
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
            Boolean b = Salva();
            if (!b)
            {
                if (MessageBox.Show("Continui l'uscita?", this.Text, MessageBoxButtons.YesNo, MessageBoxIcon.Question)
                      == DialogResult.Yes)
                    this.Close();
            }
            else this.Close();
        }

        private void FillTabs()
        {
            DataRow x;
            string p = "";
            string s = "";
            DataTable t = new DataTable();

            p = "TabPos";
            s = "SELECT * FROM " + p + " WHERE tab_ann=0";
            t = _clsFun.FillTabSql(p, s, false, _strConSql);
            x = t.NewRow();
            x["tab_cod"] = "";
            x["tab_des"] = "  Non definito";
            t.Rows.InsertAt(x, 0);
            cmbCnfPos.DataSource = t;
            cmbCnfPos.DisplayMember = "tab_des";
            cmbCnfPos.ValueMember = "tab_cod";
            cmbCnfPos.SelectedValue = "";

            p = "TabBilance";
            s = "SELECT * FROM " + p + " WHERE tab_ann=0";
            t = _clsFun.FillTabSql(p, s, false, _strConSql);
            x = t.NewRow();
            x["tab_cod"] = "";
            x["tab_des"] = "  Non definito";
            t.Rows.InsertAt(x, 0);
            cmbCnfBil.DataSource = t;
            cmbCnfBil.DisplayMember = "tab_des";
            cmbCnfBil.ValueMember = "tab_cod";
            cmbCnfBil.SelectedValue = "";

            p = "TabReparti";
            s = "SELECT * FROM " + p + " WHERE tab_ann=0";
            t = _clsFun.FillTabSql(p, s, false, _strConSql);
            x = t.NewRow();
            x["tab_cod"] = "";
            x["tab_des"] = "  Non definito";
            t.Rows.InsertAt(x, 0);
            cmbCnfRfo.DataSource = t;
            cmbCnfRfo.DisplayMember = "tab_des";
            cmbCnfRfo.ValueMember = "tab_cod";
            cmbCnfRfo.SelectedValue = "";

            //try
            //{
            //    p = "TabEtichette";
            //    s = "SELECT * FROM " + p + " WHERE tab_ann=0";
            //    t = _clsFun.FillTabSql(p, s, false, _strConSql);
            //    x = t.NewRow();
            //    x["tab_cod"] = "";
            //    x["tab_des"] = "  Non definito";
            //    t.Rows.InsertAt(x, 0);
            //    cmbCnfEti.DataSource = t;
            //    cmbCnfEti.DisplayMember = "tab_des";
            //    cmbCnfEti.ValueMember = "tab_cod";
            //    cmbCnfEti.SelectedValue = "";
            //}
            //catch (Exception ex)
            //{
            //    Console.WriteLine(ex.Message);
            //}

        }

        private void FillDati()
        {
            string s = "SELECT * FROM AnaConfigurazione WHERE cnf_cod='" + txtCnfCod.Text + "'";
            DataTable t = _clsFun.FillTabSql(TABANA, s, true, _strConSql);
            if (t.Rows.Count > 0)
            {
                txtCnfCod.Text  =  (string)t.Rows[0]["cnf_cod"];
                txtCnfRag.Text  =  (string)t.Rows[0]["cnf_rag"];
                txtCnfInd.Text  =  (string)t.Rows[0]["cnf_ind"];
                txtCnfLoc.Text  =  (string)t.Rows[0]["cnf_loc"];
                txtCnfCap.Text  =  (string)t.Rows[0]["cnf_cap"];
                txtCnfPrv.Text  =  (string)t.Rows[0]["cnf_prv"];
                txtCnfPiv.Text  =  (string)t.Rows[0]["cnf_piv"];
                txtCnfCfi.Text = (string)t.Rows[0]["cnf_cfi"];
                txtCnfBan.Text = (string)t.Rows[0]["cnf_ban"];
                txtCnfIba.Text = (string)t.Rows[0]["cnf_iba"];
                txtCnfTel.Text = (string)t.Rows[0]["cnf_tel"];
                txtCnfWeb.Text = (string)t.Rows[0]["cnf_web"];
                txtCnfMai.Text = (string)t.Rows[0]["cnf_mai"];
                txtCnfV01.Text = (string)t.Rows[0]["cnf_v01"];

                cmbCnfPos.SelectedValue = (string)t.Rows[0]["cnf_pos"];
                if (cmbCnfPos.SelectedValue == null)
                    cmbCnfPos.SelectedValue = "";

                cmbCnfBil.SelectedValue = (string)t.Rows[0]["cnf_bil"];
                if (cmbCnfBil.SelectedValue == null)
                    cmbCnfBil.SelectedValue = "";

                cmbCnfRfo.SelectedValue = (string)t.Rows[0]["cnf_rfo"];
                if (cmbCnfRfo.SelectedValue == null)
                    cmbCnfRfo.SelectedValue = "";

                //try
                //{
                //    cmbCnfEti.SelectedValue = (string)t.Rows[0]["cnf_eti"];
                //    if (cmbCnfEti.SelectedValue == null)
                //        cmbCnfEti.SelectedValue = "";
                //}
                //catch (Exception ex)
                //{
                //    Console.WriteLine(ex.Message);
                //}

                Console.WriteLine("aaaaaaaaaa");

            }
        }

        private Boolean Salva()
        {
            Boolean b = true;
            string s = "";
            DataRow y;

            if (txtCnfCod.Text =="")
            {
                MessageBox.Show("Codice mancante!");
                return false;
            }

            s = "SELECT * FROM " + TABANA + " WHERE cnf_cod = '" + txtCnfCod.Text + "'";
            DataTable t = _clsFun.FillTabSql(TABANA, s, true, _strConSql);

            y = t.NewRow();
            y["cnf_cod"] = txtCnfCod.Text;
            y["cnf_rag"] = txtCnfRag.Text;
            y["cnf_ind"] = txtCnfInd.Text;
            y["cnf_loc"] = txtCnfLoc.Text;
            y["cnf_cap"] = txtCnfCap.Text;
            y["cnf_prv"] = txtCnfPrv.Text;
            y["cnf_piv"] = txtCnfPiv.Text;
            y["cnf_cfi"] = txtCnfCfi.Text;
            y["cnf_ban"] = txtCnfBan.Text;
            y["cnf_iba"] = txtCnfIba.Text;
            y["cnf_pos"] = cmbCnfPos.SelectedValue.ToString();
            y["cnf_bil"] = cmbCnfBil.SelectedValue.ToString();
            y["cnf_rfo"] = cmbCnfRfo.SelectedValue.ToString();
            //y["cnf_eti"] = cmbCnfEti.SelectedValue.ToString();
            y["cnf_tel"] = txtCnfTel.Text;
            y["cnf_web"] =  txtCnfWeb.Text.Trim();
            y["cnf_mai"] = txtCnfMai.Text.Trim();
            y["cnf_v01"] = txtCnfV01.Text.Trim();

            if (t.Rows.Count == 0)
                s = _clsFun.SqlInsertRow(TABANA, t, y);
            else
            {
                ArrayList aWhe = new ArrayList();
                aWhe.Add("cnf_cod");
                ArrayList aExl = new ArrayList();
                s = _clsFun.SqlUpdRow(TABANA, t, t.Rows[0], y, aWhe, aExl);
            }

            if (s != "")
            {
                _clsFun.SqlWrite(s, _strConSql);
                _clsFun.FileLog(TABANA, txtCnfCod.Text, s);
            }

            return (b);
        }

        private void chkAzzArt_Click(object sender, EventArgs e)
        {
            if(chkAzzArt.Checked)
            {
                chkAzzLis.Checked = true;
                chkAzzMov.Checked = true;
                chkAzzVar.Checked = true;
                //chkAzzSta.Checked = true;
                chkAzzOff.Checked = true;
                chkAzzLog.Checked = true;
                chkAzzTdv.Checked = true;
                chkAzzOrf.Checked = true;
                chkAzzDif.Checked = true;
            }
        }

        private void chkAll_Click(object sender, EventArgs e)
        {
            foreach (Control c in panel2.Controls)
            {
                if (c is CheckBox && c.Name.Substring(0,6) == "chkAzz")
                {
                    ((CheckBox)c).Checked = chkAll.Checked;
                }
            }
        }

        private void btnAzz_Click(object sender, EventArgs e)
        {
            SqlDelete();
        }

        private void SqlDelete()
        {
            string s = "";

            if (chkAzzArt.Checked)
            {
                s = "DELETE FROM AnaArticoli";
                _clsFun.SqlWrite(s, _strConSql);
                s = "DELETE FROM AnaBarcode";
                _clsFun.SqlWrite(s, _strConSql);
                _clsFun.NumSet(_clsDef.COD04Z, clsDefine.enuNumeratori.NumAnaArticoli, "0", _strConSql);
            }
            if (chkAzzCli.Checked)
            {
                s = "DELETE FROM AnaClienti";
                _clsFun.SqlWrite(s, _strConSql);
                _clsFun.NumSet(_clsDef.COD04Z, clsDefine.enuNumeratori.NumAnaClienti, "0", _strConSql);
                s = "DELETE FROM AnaCliDestinazioni";
                _clsFun.SqlWrite(s, _strConSql);
            }
            if (chkAzzFor.Checked)
            {
                s = "DELETE FROM AnaFornitori";
                _clsFun.SqlWrite(s, _strConSql);
                _clsFun.NumSet(_clsDef.COD04Z, clsDefine.enuNumeratori.NumAnaFornitori, "0", _strConSql);
            }
            if (chkAzzTes.Checked)
            {
                s = "DELETE FROM AnaTessere";
                _clsFun.SqlWrite(s, _strConSql);
                s = "DELETE FROM AnaTessLegami";
                _clsFun.SqlWrite(s, _strConSql);
                _clsFun.NumSet(_clsDef.COD04Z, clsDefine.enuNumeratori.NumAnaTessere, "0", _strConSql);
            }
            if (chkAzzEcr.Checked)
            {
                s = "DELETE FROM TabEcrLv1";
                _clsFun.SqlWrite(s, _strConSql);
                s = "DELETE FROM TabEcrLv2";
                _clsFun.SqlWrite(s, _strConSql);
                s = "DELETE FROM TabEcrLv3";
                _clsFun.SqlWrite(s, _strConSql);
            }
            if (chkAzzRep.Checked)
            {
                s = "DELETE FROM TabReparti";
                _clsFun.SqlWrite(s, _strConSql); 
                s = "DELETE FROM TabReBilance";
                _clsFun.SqlWrite(s, _strConSql);
            }
            if (chkAzzVar.Checked)
            {
                s = "DELETE FROM GesVariazioni";
                _clsFun.SqlWrite(s, _strConSql);
            }
            if (chkAzzMov.Checked)
            {
                //s = "DELETE FROM GesMovimenti";
                //_clsFun.SqlWrite(s, _strConSql);

                int n = 1;

                s = "SELECT COUNT(*) AS cnt FROM GesMovimenti";
                DataTable t = _clsFun.FillTabSql("GesMovimenti", s, false, _strConSql);

                if (t.Rows.Count > 0)
                {
                    Int64 i = Convert.ToInt64(t.Rows[0]["cnt"]);

                    while (i > 0)
                    {
                        s = "DELETE TOP (10000) FROM GesMovimenti";
                        _clsFun.SqlWrite(s, _strConSql);

                        i -= 10000;

                        chkAzzMov.Text = "GesMovimenti " + n.ToString() + "(" + i.ToString() + ")";
                        Application.DoEvents();
                    }
                    chkAzzMov.Text = "Fat/Movimenti";
                }

                s = "DELETE FROM GesMovTestate";
                _clsFun.SqlWrite(s, _strConSql);
                s = "DELETE FROM GesFatTestate";
                _clsFun.SqlWrite(s, _strConSql);
                s = "DELETE FROM GesInventario";
                _clsFun.SqlWrite(s, _strConSql);
                s = "DELETE FROM GesInvTestate";
                _clsFun.SqlWrite(s, _strConSql);
                _clsFun.NumSet(DateTime.Today.Year.ToString(), clsDefine.enuNumeratori.NumGesMovimenti, "0", _strConSql);
                _clsFun.NumSet(DateTime.Today.Year.ToString(), clsDefine.enuNumeratori.NumGesMovFatture, "0", _strConSql);
                _clsFun.NumSet(DateTime.Today.Year.ToString(), clsDefine.enuNumeratori.NumGesDocDdt, "0", _strConSql);
                _clsFun.NumSet(DateTime.Today.Year.ToString(), clsDefine.enuNumeratori.NumGesDocFatture, "0", _strConSql);
                _clsFun.NumSet(DateTime.Today.Year.ToString(), clsDefine.enuNumeratori.NumGesFatElettroniche, "0", _strConSql);
            }
            if (chkAzzOff.Checked)
            {
                s = "DELETE FROM GesOffArticoli";
                _clsFun.SqlWrite(s, _strConSql);
                s = "DELETE FROM GesOffTestate";
                _clsFun.SqlWrite(s, _strConSql);
                _clsFun.NumSet(DateTime.Today.Year.ToString(), clsDefine.enuNumeratori.NumGesOfferte, "0", _strConSql);
            }
            if (chkAzzLis.Checked)
            {
                s = "DELETE FROM GesLisVendita";
                _clsFun.SqlWrite(s, _strConSql);
                s = "DELETE FROM GesLisAcquisto";
                _clsFun.SqlWrite(s, _strConSql);
            }
            if (chkAzzSta.Checked)
            {
                //s = "DELETE FROM GesNegMca";
                //_clsFun.SqlWrite(s, _strConSqlSta);
                //s = "DELETE FROM GesNegVet";
                //_clsFun.SqlWrite(s, _strConSqlSta);
                //s = "DELETE FROM GesNegVep";
                //_clsFun.SqlWrite(s, _strConSqlSta);


                for (int n = 1; n <= 4; n++)
                {
                    string sTab = "GesNegMca";
                    if (n == 2)
                        sTab = "GesNegVet";
                    else if (n == 3)
                        sTab = "GesNegVep";
                    else if (n == 4)
                        sTab = "GesNegVen";

                    s = "SELECT COUNT(*) AS cnt FROM " + sTab;
                    DataTable t = _clsFun.FillTabSql(sTab, s, false, _strConSqlSta);

                    if (t.Rows.Count > 0)
                    {
                        Int64 i = Convert.ToInt64(t.Rows[0]["cnt"]);

                        while (i > 0)
                        {
                            s = "DELETE TOP (10000) FROM " + sTab;
                            _clsFun.SqlWrite(s, _strConSqlSta);

                            i -= 10000;

                            chkAzzSta.Text = "Statistiche " + n.ToString() + "(" + i.ToString() + ")";
                            Application.DoEvents();
                        }
                        chkAzzSta.Text = "Statistiche";
                    }
                }
            }
            if (chkAzzSts.Checked)
            {
                s = "DELETE FROM StaPerArticoli";
                _clsFun.SqlWrite(s, _strConSqlStaSto);
                s = "DELETE FROM StaPerReparti";
                _clsFun.SqlWrite(s, _strConSqlStaSto);
            }
            if (chkAzzOrf.Checked)
            {
                s = "DELETE FROM GesOrdForRighe";
                _clsFun.SqlWrite(s, _strConSql);
                s = "DELETE FROM GesOrdForTestate";
                _clsFun.SqlWrite(s, _strConSql);
                _clsFun.NumSet(DateTime.Today.Year.ToString(), clsDefine.enuNumeratori.NumGesOrdFor, "0", _strConSql);
            }
            if (chkAzzTfo.Checked)
            {
                s = "DELETE FROM TabTabFornitori";
                _clsFun.SqlWrite(s, _strConSql);
            }
            if (chkAzzDif.Checked)
            {
                s = "DELETE FROM GesDivForArticoli";
                _clsFun.SqlWrite(s, _strConSql);
                s = "DELETE FROM GesDivForBarcode";
                _clsFun.SqlWrite(s, _strConSql);
                s = "DELETE FROM GesDivForTestate";
                _clsFun.SqlWrite(s, _strConSql);
            }
            if (chkAzzTdv.Checked)
            {
                s = "DELETE FROM TmpDiv";
                _clsFun.SqlWrite(s, _strConSql);
            }
            if (chkAzzLog.Checked)
            {
                s = "DELETE FROM LogLog";
                _clsFun.SqlWrite(s, _strConSqlLog);
            }
            if (chkAzzTra.Checked)
            {
                s = "DELETE FROM AnaForDivDettaglio";
                _clsFun.SqlWrite(s, _strConSql);
                s = "DELETE FROM AnaForDivTestate";
                _clsFun.SqlWrite(s, _strConSql);
                s = "DELETE FROM AnaForDivTipi";
                _clsFun.SqlWrite(s, _strConSql);
            }
            if (chkAzzLot.Checked)
            {
                s = "DELETE FROM GesDocLotti";
                _clsFun.SqlWrite(s, _strConSql);

                //s = "DELETE FROM GesMovBilLotti";
                //_clsFun.SqlWrite(s, _strConSql);

                _clsFun.NumSet(DateTime.Today.Year.ToString(), clsDefine.enuNumeratori.NumGesDocLotti, "0", _strConSql);
            }

            MessageBox.Show("Azzeramento completato!", "CANCELLAZIONE DATI", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void tabelleToolStripMenuItem_Click(object sender, EventArgs e)
        {
            new frmUtyImpianto().ShowDialog();
        }

        private void datBaseToolStripMenuItem_Click(object sender, EventArgs e)
        {
            new frmUtyDataBase().ShowDialog();
        }

        private void varieToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                frmWait.ShowWait("Accesso in corso...");
                new frmUtyVarie().ShowDialog(this);
            }
            finally
            {
                frmWait.CloseWait();
            }
        }

        private void etichetteToolStripMenuItem_Click(object sender, EventArgs e)
        {
        }

        private void fontsToolStripMenuItem_Click(object sender, EventArgs e)
        {
            new frmUtyGraph().ShowDialog();
        }

        private void parametriToolStripMenuItem_Click(object sender, EventArgs e)
        {
            frmGesTabelle f = new frmGesTabelle();
            f._strTab = "TabParametri";
            f._strTabDes = "Parametri";
            f._bolCodAlf = false;
            f._intCodLen = 3;
            f._intDesLen = 50;
            f._bolColAnn = true;
            f.ShowDialog();
        }

        private void utentiToolStripMenuItem_Click(object sender, EventArgs e)
        {
            frmGesTabelle f = new frmGesTabelle();
            f._strTab = "TabUtenti";
            f._strTabDes = "Utenti";
            f._bolCodAlf = false;
            f._intCodLen = 2;
            f._intDesLen = 20;
            f._bolColAnn = true;
            f.ShowDialog();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            new frmUtyConf().ShowDialog();
        }
    }
}
