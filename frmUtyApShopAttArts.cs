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
    public partial class frmUtyApShopAttArts : Form
    {
        clsDefine _clsDef = new clsDefine();
        clsFuncs _clsFun = new clsFuncs();

        private string _strConSql = "";

        public frmUtyApShopAttArts()
        {
            InitializeComponent();
        }

        private void frmUtyApShopAttArts_Load(object sender, EventArgs e)
        {
            _strConSql = _clsFun.ConSql("");
        }
        private void frmUtyApShopAttArts_KeyDown(object sender, KeyEventArgs e)
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

        private void estrazioneToolStripMenuItem_Click(object sender, EventArgs e)
        {
            FillDati();
        }

        private void btnMdb_Click(object sender, EventArgs e)
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

        private void FillDati()
        {
            string p = "";
            string s = "";
            DataRow x;
            DataRow[] j;
            DataRow[] j2;

            DataTable tTmp = new clsGenTabTmp().TabTmpArtCtrl2("TmpArt");
            DataColumn[] keys = new DataColumn[2];
            keys[0] = tTmp.Columns["tmp_art"];

            string sCn = _clsFun.ConMdb(lblPath.Text);

            p = "AnaArticoli";
            s = "SELECT ";
            s += "AnaArticoli.art_cod, ";
            s += "AnaArticoli.art_des, ";
            s += "AnaArticoli.art_sta, ";
            s += "AnaArticoli.art_iva, ";
            s += "AnaArticoli.art_umi, ";
            s += "AnaArticoli.art_rep, ";
            s += "AnaArticoli.art_pve, ";
            s += "AnaBarcode.ean_ean ";
            s += "FROM AnaArticoli INNER JOIN AnaBarcode ON AnaArticoli.art_cod = AnaBarcode.ean_art;";
            DataTable tMdb = _clsFun.FillTabMdb(p, s, false, sCn);
            keys = new DataColumn[2];
            keys[0] = tMdb.Columns["ean_ean"];
            keys[1] = tMdb.Columns["art_cod"];

            s = "SELECT ";
            s += "AnaArticoli.art_cod, ";
            s += "AnaArticoli.art_des, ";
            s += "AnaArticoli.art_sta, ";
            s += "AnaArticoli.art_umi, ";
            s += "AnaBarcode.ean_ean, ";
            s += "AnaArticoli.art_dtm ";
            s += "FROM AnaArticoli INNER JOIN AnaBarcode ON AnaArticoli.art_cod = AnaBarcode.ean_art";
            DataTable tSql = _clsFun.FillTabSql(p, s, false, _strConSql);

            progressBar1.Value = 0;
            progressBar1.Maximum = tSql.Rows.Count;
            progressBar1.Minimum = 0;

            foreach(DataRow y in tSql.Rows)
            {
                progressBar1.Increment(1);
                Application.DoEvents();

                j = tMdb.Select("ean_ean='" + (string)y["ean_ean"] + "'");
                if(j.Length > 0)
                {
                    j2 = tTmp.Select("tmp_art='" + (string)y["art_cod"] + "'");
                    if (j2.Length == 0)
                    {
                        x = tTmp.NewRow();
                        x["tmp_art"] = y["art_cod"];
                        x["tmp_des"] = y["art_des"];
                        x["tmp_sta"] = "N";
                        x["mdb_art"] = j[0]["art_cod"];
                        x["mdb_des"] = j[0]["art_des"];
                        tTmp.Rows.Add(x);
                    }
                }
            }

            dgv1.DataSource = tTmp;

            lblCnt.Text = tTmp.Rows.Count.ToString();

        }

        private void btnOk_Click(object sender, EventArgs e)
        {
            ArtAttiva();
        }

        private void ArtAttiva()
        {
            string s = "";

            DataTable t = (DataTable)dgv1.DataSource;

            progressBar1.Value = 0;
            progressBar1.Maximum = t.Rows.Count;
            progressBar1.Minimum = 0;

            foreach(DataRow y in t.Rows)
            {
                progressBar1.Increment(1);
                Application.DoEvents();

                s = "UPDATE AnaArticoli SET art_sta='A' WHERE art_cod='" + y["tmp_art"] +  "'";
                _clsFun.SqlWrite(s,_strConSql);
            }

            MessageBox.Show("Fine lavoro!");
        }

    }
}
