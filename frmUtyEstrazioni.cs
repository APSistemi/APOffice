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
    public partial class frmUtyEstrazioni : Form
    {
        clsDefine _clsDef = new clsDefine();
        clsFuncs _clsFun = new clsFuncs();
        clsVariazioni _clsVar = new clsVariazioni();
        clsQuery _clsQry = new clsQuery();

        private string _strConSql = "";

        private const string TABANAART = "AnaArticoli";
        private const string TABANAEAN = "AnaBarcode";

        public frmUtyEstrazioni()
        {
            InitializeComponent();
        }

        private void estrazioneToolStripMenuItem_Click(object sender, EventArgs e)
        {
            _strConSql = _clsFun.ConSql("");
            FillDati();
            SetDgv1();
        }

        private void esciToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Esci();
        }
        private void Esci()
        {
            this.Close();
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
            cTbc.DataPropertyName = "art_sta";
            cTbc.Name = "Stato";
            cTbc.Width = 50;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = true;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "art_cod";
            cTbc.Name = "Codice";
            cTbc.Width = 60;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = true;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "art_des";
            cTbc.Name = "Descrizione";
            cTbc.Width = 250;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = true;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "tmp_cos";
            cTbc.Name = "Costo";
            cTbc.Width = 50;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = false;
            dgv1.Columns.Add(cTbc);
            dgv1.Columns[cTbc.Name].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dgv1.Columns[cTbc.Name].DefaultCellStyle.Format = "##0.000";

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "tmp_prv";
            cTbc.Name = "Prezzo vendita";
            cTbc.Width = 50;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = false;
            dgv1.Columns.Add(cTbc);
            dgv1.Columns[cTbc.Name].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dgv1.Columns[cTbc.Name].DefaultCellStyle.Format = "##0.00";

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "tmp_msg";
            cTbc.Name = "Fornitori";
            cTbc.Width = 250;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = true;
            dgv1.Columns.Add(cTbc);
        }

        private void FillDati()
        {
            string s = "SELECT art_sta, art_cod, art_des, art_umi, art_iva, art_tgr, art_pxc, art_pne, art_reb, art_plu FROM " + TABANAART + " ";
            if(chkRep.Checked)
                s += "WHERE art_sta='" + _clsDef.STAREP + "' ";
            else if (chkAtt.Checked)
                s += "WHERE art_sta='" + _clsDef.STAATT + "' ";
            s += "ORDER BY art_cod";
            DataTable t = _clsFun.FillTabSql(TABANAART, s, false, _strConSql);
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.Decimal"),
                ColumnName = "tmp_cos",
                Caption = "Costo",
                ReadOnly = false,
                DefaultValue = (Decimal)0
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.Decimal"),
                ColumnName = "tmp_prv",
                Caption = "Vendita",
                ReadOnly = false,
                DefaultValue = (Decimal)0
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "tmp_ean",
                Caption = "Barcode",
                MaxLength = 13,
                ReadOnly = false,
                DefaultValue = (String)""
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "tmp_msg",
                Caption = "Note",
                MaxLength = 200,
                ReadOnly = false,
                DefaultValue = (String)""
            });

            progressBar1.Value = 0;
            progressBar1.Visible = true;
            progressBar1.Maximum = t.Rows.Count;
            progressBar1.Minimum = 0;

            foreach (DataRow y in t.Rows)
            {
                progressBar1.Increment(1);
                Application.DoEvents();

                //if (progressBar1.Value > 1000)
                //    break;

                DataTable tTmp = _clsQry.ArtSeek((string)y["art_cod"], "");
                if (tTmp.Rows.Count > 0)
                {
                    y["tmp_cos"] = tTmp.Rows[0]["tmp_cos"];
                    y["tmp_prv"] = tTmp.Rows[0]["tmp_prv"];
                }

                tTmp = _clsQry.ArtFornitori((string)y["art_cod"]);

                if (tTmp.Rows.Count > 0)
                {
                    foreach (DataRow x in tTmp.Rows)
                    {
                        y["tmp_msg"] += (string)x["for_des"] + " - ";
                    }
                }
            }

            dgv1.DataSource = t;

            lblCnt.Text = t.Rows.Count.ToString("###,##0");
        }

        private void generazioneToolStripMenuItem_Click(object sender, EventArgs e)
        {
            GenKonzOpen();
        }

        private void GenKonzOpen()
        {
            string s = "";
            DataRow[] j;
            string sFil = "C:\\ApProject\\KONZOpen.txt";

            StreamWriter sw = new StreamWriter(sFil, false);

            s = "SELECT ean_art, ean_ean FROM AnaBarcode ";
            s += "LEFT JOIN AnaArticoli ON AnaBarcode.ean_art = AnaArticoli.art_cod ";
            s += "WHERE AnaArticoli.art_sta='A' OR AnaArticoli.art_sta='R'";
            DataTable tEan = _clsFun.FillTabSql(TABANAEAN, s, false, _strConSql);
            DataColumn[] Key = new DataColumn[2] 
            { 
                tEan.Columns["ean_art"] ,
                tEan.Columns["ean_ean"] 
            };
            tEan.PrimaryKey = Key;

            DataTable t = (DataTable)dgv1.DataSource;

            string sRig = "";

            foreach(DataRow y in t.Rows)
            {

                j = tEan.Select("ean_art='" + y["art_cod"] + "'");

                if (j.Length > 0)
                {
                    for (int i = 0; i < j.Length; i++)
                    {
                        sRig = new string(' ', 6);                                          //Cod cliente
                        sRig += new string(' ', 6);                                         //Numero fattura

                        s = DateTime.Now.ToString("ddMMyyyy");
                        sRig += s;                                                          //Data del giorno
                        //sRig += ((string)y["art_cod"]).Substring(1, 6);                     //Codice articolo
                        sRig += new string(' ', 6);                                          //Cod articolo

                        s = (string)y["art_des"] + new string(' ', 30);
                        sRig += s.Substring(0,30);                                         //Descrizione

                        sRig += "   ";                                                      //Filler

                        s = "  ";
                        if(((string)y["art_tgr"]).Length == 2)
                            s = ((string)y["art_tgr"]).Substring(0, 2);
                        sRig += s;                                                          //Tipo grammatura

                        sRig += "        ";                                                 //ECR

                        s = ((string)y["art_iva"]).Trim();
                        if (s.Length != 3)
                            s = "   ";
                        sRig += s.Substring(1, 2);                                          //IVA

                        s = "  ";
                        if (((string)y["art_umi"]).Trim().Length == 2)
                            s = (string)y["art_umi"];
                        sRig += s;                                                          //UMI

                        sRig += Convert.ToInt32(y["art_pxc"]).ToString("0000");             //PXC
                        sRig += ((decimal)y["art_pne"]).ToString("000000");                 //Grammatura
                        sRig += " ";                                                        //Nota accredito
                        sRig += "0000000";                                                  //Qta fattura
                        sRig += ((decimal)y["tmp_cos"] * 100).ToString("000000000");        //Costo
                        sRig += "00000000000";                                              //Importo riga
                        sRig += new string(' ', 16);                                        //EAN vuoto
                        sRig += "A";                                                        //Tipo riga
                        sRig += "0000000000000";                                            //Importo unitario
                        sRig += ((string)j[i]["ean_ean"]).PadRight(16, Convert.ToChar(' ')); //EAN
                        sRig += ((decimal)y["tmp_prv"] * 100).ToString("00000000000");      //Prezzo pubblico
                        sRig += new string(' ', 61);                                        //Filler
                        sRig += ((string)y["art_reb"] + " ").Substring(0, 1);               //Reparto bilancia
                        s = "   ";
                        if (((string)y["art_plu"]).Trim() != "")
                            s = ((string)y["art_plu"]).Substring(1, 3);
                        sRig += s;                                                          //PLU
                        if(s.Trim() != "")
                            sRig += "P";                                                        //Corpo / peso
                        else
                            sRig += " ";                                                        //Corpo / peso
                        sRig += "A"+((string)y["art_cod"]).Substring(0, 7);                 //Codice articolo
                        sRig += new string(' ', 86);                                        //Filler
                        sw.Write(sRig + _clsDef.CRLF);
                    }
                }
            }


            ((TextWriter)sw).Flush();
            sw.Close();
            sw.Dispose();

        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {

        }

    }
}
