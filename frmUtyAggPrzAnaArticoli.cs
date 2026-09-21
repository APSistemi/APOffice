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
    public partial class frmUtyAggPrzAnaArticoli : Form
    {
        clsDefine _clsDef = new clsDefine();
        clsFuncs _clsFun = new clsFuncs();

        private string _strConSql = "";

        public frmUtyAggPrzAnaArticoli()
        {
            InitializeComponent();
            new clsGesGraph().SetGraph(this, 0);
            _strConSql = _clsFun.ConSql("");
            //FillTabs();
            //SetDgv1();

        }

        private void frmUtyAggPrzAnaArticoli_Load(object sender, EventArgs e)
        {

        }

        private void esciToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Esci();
        }

        private void frmUtyAggPrzAnaArticoli_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Escape)
                Esci();
        }

        private void Esci()
        {
            this.Close();
        }

        private void eseguiToolStripMenuItem_Click(object sender, EventArgs e)
        {
            FillDati();
        }

        private void FillDati()
        {
            DataRow[] j;

            string s = "SELECT art_cod, art_des, art_cos, art_prv FROM AnaArticoli ORDER BY art_des";
            DataTable tArt = _clsFun.FillTabSql("AnaArticoli", s, false, _strConSql);
            tArt.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "ArtFl1",
                Caption = "Fl1",
                MaxLength = 1,
                ReadOnly = false,
                DefaultValue = (String)""
            });

            s = "SELECT ";
            s += "GesLisAcquisto_1.lia_art, ";
            s += "GesLisAcquisto_1.lia_cos ";
            //s += "AnaArticoli.art_des ";
            s += "FROM ";
            s += "(SELECT lia_tip, lia_for, MAX(lia_dti) AS lia_dti, lia_art ";
            s += " FROM GesLisAcquisto ";
            s += " GROUP BY lia_tip, lia_for, lia_art) AS A ";
            s += "LEFT OUTER JOIN ";
            s += "GesLisAcquisto AS GesLisAcquisto_1 ON A.lia_tip = GesLisAcquisto_1.lia_tip AND A.lia_for = GesLisAcquisto_1.lia_for AND A.lia_art = GesLisAcquisto_1.lia_art ";
            //s += "LEFT OUTER JOIN AnaArticoli ON AnaArticoli.art_cod = A.lia_art ";
            DataTable tLia = _clsFun.FillTabSql("Lia", s, false, _strConSql);
            DataColumn[] keys = new DataColumn[1];
            keys[0] = tLia.Columns["art_cod"];
            tLia.PrimaryKey = keys;

            //s = "SELECT ";
            //s += "GesLisVendita_1.liv_lis, ";
            //s += "GesLisVendita_1.liv_art, ";
            //s += "GesLisVendita_1.liv_prv ";
            //s += "FROM (";
            //s += "SELECT liv_lis, MAX(liv_dti) AS liv_dti, liv_art ";
            //s += "FROM  GesLisVendita ";
            //s += "WHERE liv_ann=0 ";
            //s += "GROUP BY liv_lis, liv_art) AS A ";
            //s += "LEFT OUTER JOIN GesLisVendita AS GesLisVendita_1 ON ";
            //s += "A.liv_lis = GesLisVendita_1.liv_lis AND ";
            //s += "A.liv_art = GesLisVendita_1.liv_art ";

            s = "SELECT * FROM (";
            s += "SELECT ";
            s += "ROW_NUMBER() OVER (PARTITION BY liv_lis, liv_art ORDER BY liv_lis, liv_art, liv_dti DESC) AS ROW, ";
            s += "liv_lis, ";
            s += "liv_art, ";
            s += "liv_prv ";
            s += "FROM GesLisVendita ) AS A WHERE ROW = 1";
            DataTable tLiv = _clsFun.FillTabSql("TabVen", s, false, _strConSql);
            keys = new DataColumn[2];
            keys[0] = tLiv.Columns["liv_lis"];
            keys[1] = tLiv.Columns["liv_art"];
            tLiv.PrimaryKey = keys;

            Console.WriteLine("aaaaaaaaa");

            progressBar1.Value = 0;
            progressBar1.Maximum = tArt.Rows.Count;
            progressBar1.Minimum = 0;

            foreach(DataRow y in tArt.Rows)
            {
                progressBar1.Increment(1);
                Application.DoEvents();

                j = tLia.Select("lia_art='" + y["art_cod"] + "'");
                if (j.Length > 0)
                {
                    if (Convert.ToDecimal(y["art_cos"]) != Convert.ToDecimal(j[0]["lia_cos"]))
                    {
                        y["art_cos"] = j[0]["lia_cos"];
                        y["ArtFl1"] = "S";
                    }
                }

                j = tLiv.Select("liv_lis='001' AND liv_art='" + y["art_cod"] + "'");
                if (j.Length > 0)
                {
                    if (Convert.ToDecimal(y["art_prv"]) != Convert.ToDecimal(j[0]["liv_prv"]))
                    {
                        y["art_prv"] = j[0]["liv_prv"];
                        y["ArtFl1"] = "S";
                    }
                }
                else
                {
                    j = tLiv.Select("liv_lis='002' AND liv_art='" + y["art_cod"] + "'");
                    if (j.Length > 0)
                    {
                        if (Convert.ToDecimal(y["art_prv"]) != Convert.ToDecimal(j[0]["liv_prv"]))
                        {
                            y["art_prv"] = j[0]["liv_prv"];
                            y["ArtFl1"] = "S";
                        }
                    }
                }
            }

            dgv1.DataSource = tArt;
        }

        private void aggiornaToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Salva();
        }

        private void Salva()
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

                if((string)y["ArtFl1"] == "S")
                {
                    s = "UPDATE AnaArticoli SET ";
                    s += "art_cos=" + ((decimal)y["art_cos"]).ToString().Replace(",",".") + ",";
                    s += "art_prv=" + ((decimal)y["art_prv"]).ToString().Replace(",", ".") + " ";
                    s += "WHERE art_cod='" + (string)y["art_cod"] + "'";
                    _clsFun.SqlWrite(s, _strConSql);
                }
            }

        }

    }
}
