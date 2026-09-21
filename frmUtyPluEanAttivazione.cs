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
    public partial class frmUtyPluEanAttivazione : Form
    {
        clsDefine _clsDef = new clsDefine();
        clsFuncs _clsFun = new clsFuncs();

        private string _strConSql = "";

        public frmUtyPluEanAttivazione()
        {
            InitializeComponent();
        }

        private void frmUtyPluEanAttivazione_Load(object sender, EventArgs e)
        {
            _strConSql = _clsFun.ConSql("");

        }

        private void esciToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void dgv1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            FillDati();
        }

        private void FillDati()
        {

            DataRow[] j;

            /* 201801001 Per barcode che iniziano per 2 e lunghi 7 aggiunge gli zeri alla fine perche a peso: ZEUSLAB */

            string s = "SELECT ean_art, ean_ean FROM AnaBarcode WHERE (SUBSTRING(ean_ean, 1, 2) > '21') ";
            DataTable tEan = _clsFun.FillTabSql("AnaBarcode", s, false, _strConSql);

            DataColumn[] keys = new DataColumn[1];
            keys[0] = tEan.Columns["ean_ean"];
            tEan.PrimaryKey = keys;



            //s = "SELECT art_sta, art_cod, art_des, art_plu, ean_ean FROM AnaArticoli ";
            //s += "LEFT OUTER JOIN AnaBarcode ON AnaBarcode.ean_art=AnaArticoli.art_cod ";
            ////s += "WHERE art_plu<>'' ";
            //s += "WHERE LEN(ean_ean) = 13 AND SUBSTRING(ean_ean,1,1) = '2' AND art_sta ='N'";

            s = "SELECT ean_art, ean_ean, RTRIM(ean_ean) + '000000' AS eanOk ";
            s += "FROM  AnaBarcode ";
            s += "WHERE (LEN(ean_ean) = 7) AND (SUBSTRING(ean_ean, 1, 1) = '2') AND (SUBSTRING(ean_ean, 2, 1) > '1')";
            DataTable t = _clsFun.FillTabSql("PLU", s, false, _strConSql);
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "tmp_fl1",
                Caption = "Fl1",
                MaxLength = 1,
                ReadOnly = false,
                DefaultValue = (String)""
            });

            int i = 0;

            foreach(DataRow y in t.Rows)
            {
                //if (((string)y["ean_ean"]).Length == 13 && ((string)y["ean_ean"]).Substring(2, 4) == (string)y["art_plu"])
                //{
                //    y["tmp_fl1"] = "S";
                //    i++;
                //}

                j = tEan.Select("ean_ean='" + y["eanOk"] + "'");
                if(j.Length > 0)
                    y["tmp_fl1"] = "S";

            }

            dgv1.DataSource = t;

            label1.Text = "Trovati " + t.Rows.Count.ToString();

        }

        private void button2_Click(object sender, EventArgs e)
        {
            Salva();
        }

        private void Salva()
        {
            string s = "";


            DataTable t = (DataTable)dgv1.DataSource;


            foreach(DataRow y in t.Rows)
            {
                if((string)y["tmp_fl1"] != "S")
                {
                    s = "UPDATE AnaBarcode SET ean_ean='" + y["eanok"] + "' WHERE ean_ean='" + y["ean_ean"] + "'";
                    //s = "UPDATE AnaArticoli SET art_sta='A' WHERE art_cod='" + y["art_cod"] + "'";
                    _clsFun.SqlWrite(s, _strConSql);

                }
            }

            MessageBox.Show("Fine");

        }

    }
}
