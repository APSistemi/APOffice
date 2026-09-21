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
    public partial class frmUtyDataBase : Form
    {
        clsDefine _clsDef = new clsDefine();
        clsFuncs _clsFun = new clsFuncs();

        private string _strConSql = "";

        private string sBtn1 = "Generazione tabelle e chiavi mancanti";
        private string sBtn5 = "Confronta con CSV/APOffice";

        public frmUtyDataBase()
        {
            InitializeComponent();
        }

        private void frmUtyDataBase_Load(object sender, EventArgs e)
        {
            _strConSql = _clsFun.ConSql("");

            cmbMdf.Items.Add("APOffice");
            cmbMdf.Items.Add("APStat");
            cmbMdf.SelectedIndex = 0;
            button1.Text = sBtn1 + cmbMdf.SelectedItem.ToString();
            button5.Text = sBtn5 + cmbMdf.SelectedItem.ToString();

            dgv1.DataSource = new clsGenTabTmp().TabTmpMdfCsv("TabTes");
            dgv2.DataSource = new clsGenTabTmp().TabTmpMdfCsv("TabTab");
            dgv3.DataSource = new clsGenTabTmp().TabTmpMdfCsv("IdxTes");
            dgv4.DataSource = new clsGenTabTmp().TabTmpMdfCsv("IdxIdx");
        }

        private void esciToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            FillDb();
        }

        private void cmbMdf_SelectionChangeCommitted(object sender, EventArgs e)
        {
            cmbMdf.Text = cmbMdf.SelectedItem.ToString();
            ConSql();
        }

        private void cmbMdf_Leave(object sender, EventArgs e)
        {
            ConSql();
        }

        private void ConSql()
        {
            int i = _strConSql.IndexOf("Catalog=");
            i += 8;
            string s = _strConSql.Substring(i);
            i = s.IndexOf(";");
            s = s.Substring(0, i);

            string sMdf = cmbMdf.Text;        // SelectedItem.ToString();

            if (s != sMdf)
            {
                _strConSql = _strConSql.Replace(s, sMdf);
                button1.Text = sBtn1 +" "+ sMdf;
                button5.Text = sBtn5 +" "+ sMdf;

                if (!_clsFun.TestMdf(_strConSql))
                    MessageBox.Show("DataBase MDF non apribile!");
            }
        }

        private void FillDb()
        {
            string s = "SELECT DISTINCT TABLE_NAME AS tab_tab FROM INFORMATION_SCHEMA.COLUMNS"; // WHERE TABLE_NAME LIKE 'TEST%'";
            DataTable t = _clsFun.FillTabSql("TmpTab", s, false, _strConSql);
            dgv1.DataSource = t;

            s = "SELECT ";
            //s += "TABLE_CATALOG, ";
            s += "TABLE_NAME AS tab_tab, ";
            s += "COLUMN_NAME AS tab_col, ";
            //s += "IS_NULLABLE, ";
            s += "DATA_TYPE AS tab_typ, ";
            s += "CHARACTER_MAXIMUM_LENGTH AS tab_len, ";
            s += "NUMERIC_PRECISION AS tab_lun, ";
            s += "NUMERIC_SCALE AS tab_dec, ";
            s += "COLUMNPROPERTY(OBJECT_ID(TABLE_NAME), COLUMN_NAME, 'IsIdentity') AS tab_idn ";
            s += "FROM INFORMATION_SCHEMA.COLUMNS "; // WHERE TABLE_NAME LIKE 'TEST%'";
            s += "ORDER BY TABLE_NAME";
            t = _clsFun.FillTabSql("tSchema", s, false, _strConSql);
            dgv2.DataSource = t;

            //            SELECT        t.name AS TableName, ind.name AS IndexName, ind.index_id AS IndexId, ic.index_column_id AS ColumnId, col.name AS ColumnName, ind.object_id, ind.name, ind.index_id, ind.type, ind.type_desc, 
            //                         ind.is_unique, ind.data_space_id, ind.ignore_dup_key, ind.is_primary_key, ind.is_unique_constraint, ind.fill_factor, ind.is_padded, ind.is_disabled, ind.is_hypothetical, ind.allow_row_locks, 
            //                         ind.allow_page_locks, ind.has_filter, ind.filter_definition, ic.object_id AS Expr1, ic.index_id AS Expr2, ic.index_column_id, ic.column_id, ic.key_ordinal, ic.partition_ordinal, ic.is_descending_key, 
            //                         ic.is_included_column, col.object_id AS Expr3, col.name AS Expr4, col.column_id AS Expr5, col.system_type_id, col.user_type_id, col.max_length, col.precision, col.scale, col.collation_name, col.is_nullable, 
            //                         col.is_ansi_padded, col.is_rowguidcol, col.is_identity, col.is_computed, col.is_filestream, col.is_replicated, col.is_non_sql_subscribed, col.is_merge_published, col.is_dts_replicated, col.is_xml_document, 
            //                         col.xml_collection_id, col.default_object_id, col.rule_object_id, col.is_sparse, col.is_column_set
            //FROM            sys.indexes AS ind INNER JOIN
            //                         sys.index_columns AS ic ON ind.object_id = ic.object_id AND ind.index_id = ic.index_id INNER JOIN
            //                         sys.columns AS col ON ic.object_id = col.object_id AND ic.column_id = col.column_id INNER JOIN
            //                         sys.tables AS t ON ind.object_id = t.object_id
            //ORDER BY TableName, IndexId, ColumnId

            s = "SELECT DISTINCT ";
            s += "t.name AS idx_tab, ";
            s += "ind.name AS idx_idx ";
            s += "FROM sys.indexes AS ind INNER JOIN ";
            s += "sys.index_columns AS ic ON ind.object_id = ic.object_id AND ind.index_id = ic.index_id INNER JOIN ";
            s += "sys.columns AS col ON ic.object_id = col.object_id AND ic.column_id = col.column_id INNER JOIN ";
            s += "sys.tables AS t ON ind.object_id = t.object_id ";
            s += "ORDER BY idx_tab, idx_idx ";
            t = _clsFun.FillTabSql("tIndex", s, false, _strConSql);
            dgv3.DataSource = t;

            s = "SELECT ";
            s += "t.name AS idx_tab, ";
            s += "ind.name AS idx_idx, ";
            s += "ind.index_id AS idx_xid, ";
            s += "col.name AS idx_col ";
            //s += "ind.name ";
            s += "FROM sys.indexes AS ind INNER JOIN ";
            s += "sys.index_columns AS ic ON ind.object_id = ic.object_id AND ind.index_id = ic.index_id INNER JOIN ";
            s += "sys.columns AS col ON ic.object_id = col.object_id AND ic.column_id = col.column_id INNER JOIN ";
            s += "sys.tables AS t ON ind.object_id = t.object_id ";
            s += "ORDER BY idx_tab, idx_idx, idx_xid";
            t = _clsFun.FillTabSql("tIndex", s, false, _strConSql);
            dgv4.DataSource = t;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (dgv1.DataSource == null || ((DataTable)dgv1.DataSource).Rows.Count == 0)
                MessageBox.Show("Dati non caricati!");
            else
                CreaDb();
        }

        private void CreaDb()
        {
            string s = "";

            DataTable tTab = (DataTable)dgv1.DataSource;
            DataTable tFld = (DataTable)dgv2.DataSource;
            DataTable tIdx = (DataTable)dgv3.DataSource;
            DataTable tIdr = (DataTable)dgv4.DataSource;

            DataRow[] j;
            DataRow[] j2;

            progressBar1.Value = 0;
            progressBar1.Visible = true;
            progressBar1.Maximum = tTab.Rows.Count;
            progressBar1.Minimum = 0;

            foreach (DataRow y in tTab.Rows)
            {
                progressBar1.Increment(1);
                Application.DoEvents();

                string sTab = (string)y["tab_tab"];
                string sSql = "";
                string sKey = "";

                if (sTab.Substring(0, 3).ToLower() != "old")
                {
                    j = tFld.Select("tab_tab='" + sTab + "'");
                    if (j.Length > 0)
                    {
                        sSql = "CREATE TABLE " + sTab + "(";

                        for (int i = 0; i < j.Length; i++)
                        {
                            sSql += (string)j[i]["tab_col"] + " ";
                            if (Convert.ToInt16(j[i]["tab_idn"]) == 1)
                                sSql += "[int] IDENTITY(1,1) NOT NULL, ";
                            else
                            {
                                sSql += (string)j[i]["tab_typ"];
                                if (Convert.ToDecimal(j[i]["tab_len"]) > 0)
                                    sSql += "(" + Convert.ToString(j[i]["tab_len"]) + ")";
                                else if ((string)j[i]["tab_typ"] != "int" && Convert.ToDecimal(j[i]["tab_lun"]) > 0)
                                    sSql += "(" + Convert.ToString(j[i]["tab_lun"]) + "," + Convert.ToString(j[i]["tab_dec"]) + ")";

                                sSql += " NOT NULL, ";
                            }
                        }

                        sSql = sSql.Substring(0, sSql.Length - 2);
                        sSql += ")";

                        _clsFun.SqlWrite(sSql, _strConSql);
                    }

                    j = tIdx.Select("idx_tab='" + sTab + "'");
                    if (j.Length > 0)
                    {
                        for (int i = 0; i < j.Length; i++)
                        {
                            j2 = tIdr.Select("idx_tab='" + j[i]["idx_tab"] + "' AND idx_idx='" + j[i]["idx_idx"] + "'");

                            if (j2.Length > 0)
                            {
                                if (((string)j[i]["idx_idx"]).Substring(0, 2).ToLower() == "pk")
                                    sSql = "ALTER TABLE " + j[i]["idx_tab"] + " ADD CONSTRAINT " + (string)j[i]["idx_idx"] + " PRIMARY KEY (";
                                else
                                    sSql = "CREATE UNIQUE INDEX " + j[i]["idx_idx"] + " ON " + j[i]["idx_tab"] + " (";

                                for (int i2 = 0; i2 < j2.Length; i2++)
                                {
                                    sSql += (string)j2[i2]["idx_col"] + ", ";
                                }
                                sSql = sSql.Substring(0, sSql.Length - 2);
                                sSql += ")";

                                _clsFun.SqlWrite(sSql, _strConSql);
                            }
                        }
                    }

                }
            }

            lblMsg.Text += "Create tabelle ";
        }

        private void button3_Click(object sender, EventArgs e)
        {
            if (dgv1.DataSource == null || ((DataTable)dgv1.DataSource).Rows.Count == 0)
                MessageBox.Show("Dati non caricati!");
            else
                GenCsv();
        }

        private void GenCsv()
        {
            DataTable tFld = (DataTable)dgv2.DataSource;
            DataTable tIdr = (DataTable)dgv4.DataSource;

            string sFil = _clsDef.PATHTABCSV;

            StreamWriter sw = new StreamWriter(sFil, false);
            //string s = "";

            progressBar1.Value = 0;
            progressBar1.Visible = true;
            progressBar1.Maximum = tFld.Rows.Count;
            progressBar1.Minimum = 0;

            foreach (DataRow y in tFld.Rows)
            {
                progressBar1.Increment(1);
                Application.DoEvents();

                string sRig = "";
                foreach (DataColumn c in tFld.Columns)
                {
                    sRig += Convert.ToString(y[c.ToString()]).Trim() + ";";
                }
                sw.Write("TAB;" + sRig + "\r\n");
            }

            lblMsg.Text += tFld + ": generato ";

            //((TextWriter)sw).Flush();
            //sw.Close();
            //sw.Dispose();

            progressBar1.Value = 0;
            progressBar1.Visible = true;
            progressBar1.Maximum = tIdr.Rows.Count;
            progressBar1.Minimum = 0;

            foreach (DataRow y in tIdr.Rows)
            {
                progressBar1.Increment(1);
                Application.DoEvents();

                string sRig = "";
                foreach (DataColumn c in tIdr.Columns)
                {
                    sRig += Convert.ToString(y[c.ToString()]).Trim() + ";";
                }
                sw.Write("KEY;" + sRig + "\r\n");
            }

            lblMsg.Text += tIdr + ": generato ";

            ((TextWriter)sw).Flush();
            sw.Close();
            sw.Dispose();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            if (!File.Exists(_clsDef.PATHTABCSV))
                MessageBox.Show(_clsDef.PATHTABCSV + " non trovato!");
            else
                FillCsv(_clsDef.PATHTABCSV);
        }

        private void FillCsv(string strFil)
        {
            DataRow[] j;
            DataRow x;

            dgv1.DataSource = new clsGenTabTmp().TabTmpMdfCsv("TabTes");
            dgv2.DataSource = new clsGenTabTmp().TabTmpMdfCsv("TabTab");
            dgv3.DataSource = new clsGenTabTmp().TabTmpMdfCsv("IdxTes");
            dgv4.DataSource = new clsGenTabTmp().TabTmpMdfCsv("IdxIdx");

            DataTable tTab = (DataTable)dgv1.DataSource;
            DataTable tFld = (DataTable)dgv2.DataSource;
            DataTable tIdx = (DataTable)dgv3.DataSource;
            DataTable tIdr = (DataTable)dgv4.DataSource;

            using (StreamReader sr = new StreamReader(strFil))
            {
                string sRig = "";

                while ((sRig = sr.ReadLine()) != null)
                {
                    if (sRig.Length > 10)
                    {
                        string[] a = sRig.Split(new char[1] { ';' });

                        string sTip = (string)a[0];

                        if(sTip == "TAB")
                        {
                            string sTab = (string)a[1];
                            string sCol = (string)a[2];
                            string sTyp = (string)a[3];
                            int dLen = Convert.ToInt16(a[4]);
                            int dLun = Convert.ToInt16(a[5]);
                            int dDec = Convert.ToInt16(a[6]);
                            int dIdn = Convert.ToInt16(a[7]);

                            j = tTab.Select("tab_tab='" + sTab + "'");
                            if(j.Length == 0)
                            {
                                x = tTab.NewRow();
                                x["tab_tab"] = sTab;
                                tTab.Rows.Add(x);
                            }

                            x = tFld.NewRow();
                            x["tab_tab"] = sTab;
                            x["tab_col"] = sCol;
                            x["tab_typ"] = sTyp;
                            x["tab_len"] = dLen;
                            x["tab_lun"] = dLun;
                            x["tab_dec"] = dDec;
                            x["tab_idn"] = dIdn;
                            tFld.Rows.Add(x);
                        }
                        if(sTip == "KEY")
                        {
                            string sTab = (string)a[1];
                            string sIdx = (string)a[2];
                            int dXid = Convert.ToInt16(a[3]);
                            string sCol = (string)a[4];

                            j = tIdx.Select("idx_tab='" + sTab + "' AND idx_idx='" + sIdx + "'");
                            if(j.Length == 0)
                            {
                                x = tIdx.NewRow();
                                x["idx_tab"] = sTab;
                                x["idx_idx"] = sIdx;
                                tIdx.Rows.Add(x);
                            }

                            x = tIdr.NewRow();
                            x["idx_tab"] = sTab;
                            x["idx_col"] = sCol;
                            x["idx_idx"] = sIdx;
                            x["idx_xid"] = dXid;
                            tIdr.Rows.Add(x);
                        }
                    }
                }
            }
        }

        private void button5_Click(object sender, EventArgs e)
        {
            DiffCsvMdf();
        }

        private void DiffCsvMdf()
        {
            string s = "";
            DataRow[] j;
            DataRow x;
            Boolean b = false;

            DataTable tDif = new clsGenTabTmp().TabTmpMdfCsvDiff("TabDif");

            FillDb();

            DataTable tMdfTab = (DataTable)dgv2.DataSource;
            DataTable tMdfIdx = (DataTable)dgv4.DataSource;

            FillCsv(_clsDef.PATHTABCSV);

            DataTable tCsvTab = (DataTable)dgv2.DataSource;
            DataTable tCsvIdx = (DataTable)dgv4.DataSource;

            foreach(DataRow y in tCsvTab.Rows)
            {
                j = tMdfTab.Select("tab_tab='" + (string)y["tab_tab"] + "' AND tab_col='" + (string)y["tab_col"] + "'");

                if(j.Length == 0)
                {
                    x = tDif.NewRow();
                    x["dif_tip"] = "TAB";
                    x["dif_tab"] = y["tab_tab"];
                    x["dif_col"] = y["tab_col"];
                    x["csv_typ"] = y["tab_typ"];
                    x["csv_len"] = y["tab_len"];
                    x["csv_lun"] = y["tab_lun"];
                    x["csv_dec"] = y["tab_dec"];
                    x["csv_idn"] = y["tab_idn"];
                    tDif.Rows.Add(x);
                }
                else
                {
                    b = false;
                    x = tDif.NewRow();
                    x["dif_tip"] = "TAB";
                    x["dif_tab"] = (string)y["tab_tab"];
                    x["dif_col"] = (string)y["tab_col"];

                    if ((string)y["tab_typ"] != (string)j[0]["tab_typ"])
                    {
                        b = true;
                        x["csv_typ"] = y["tab_typ"];
                        x["mdf_typ"] = j[0]["tab_typ"];
                    }
                    if (Convert.ToInt32(y["tab_len"]) != Convert.ToInt32(j[0]["tab_len"]))
                    {
                        b = true;
                        x["csv_len"] = y["tab_len"];
                        x["mdf_len"] = j[0]["tab_len"];
                    }
                    if (Convert.ToInt32(y["tab_lun"]) != Convert.ToInt32(j[0]["tab_lun"]))
                    {
                        b = true;
                        x["csv_lun"] = Convert.ToDecimal(y["tab_lun"]);
                        x["mdf_lun"] = Convert.ToDecimal(j[0]["tab_lun"]);
                    }
                    if (Convert.ToInt32(y["tab_dec"]) != Convert.ToInt32(j[0]["tab_dec"]))
                    {
                        b = true;
                        x["csv_dec"] = Convert.ToDecimal(y["tab_dec"]);
                        x["mdf_dec"] = Convert.ToDecimal(j[0]["tab_dec"]);
                    }
                    if (Convert.ToInt32(y["tab_idn"]) != Convert.ToInt32(j[0]["tab_idn"]))
                    {
                        b = true;
                        x["csv_idn"] = Convert.ToDecimal(y["tab_idn"]);
                        x["mdf_idn"] = Convert.ToDecimal(j[0]["tab_idn"]);
                    }
                    if(b)
                        tDif.Rows.Add(x);
                }
            }

            foreach (DataRow y in tMdfTab.Rows)
            {
                s = (string)y["tab_col"];

                j = tCsvTab.Select("tab_tab='" + (string)y["tab_tab"] + "' AND tab_col='" + (string)y["tab_col"] + "'");

                if (j.Length == 0)
                {
                    x = tDif.NewRow();
                    x["dif_tip"] = "TAB";
                    x["dif_tab"] = y["tab_tab"];
                    x["dif_col"] = y["tab_col"];
                    x["csv_typ"] = y["tab_typ"];
                    x["csv_len"] = y["tab_len"];
                    x["csv_lun"] = y["tab_lun"];
                    x["csv_dec"] = y["tab_dec"];
                    x["csv_idn"] = y["tab_idn"];
                    tDif.Rows.Add(x);
                }
            }

            /*** INDICI ***/

            foreach (DataRow y in tCsvIdx.Rows)
            {
                j = tMdfIdx.Select("idx_tab='" + (string)y["idx_tab"] + "' AND idx_idx='" + (string)y["idx_idx"] + "'");

                if (j.Length == 0)
                {
                    x = tDif.NewRow();
                    x["dif_tip"] = "IDX";
                    x["dif_tab"] = y["idx_tab"];
                    x["dif_col"] = y["idx_idx"];
                    x["csv_typ"] = y["idx_col"];
                    x["csv_len"] = y["idx_xid"];
                    tDif.Rows.Add(x);
                }
                //else
                //{
                //    b = false;
                //    x = tDif.NewRow();
                //    x["dif_tip"] = "TAB";
                //    x["dif_tab"] = (string)y["tab_tab"];
                //    x["dif_col"] = (string)y["tab_col"];

                //    if ((string)y["tab_typ"] != (string)j[0]["tab_typ"])
                //    {
                //        b = true;
                //        x["csv_typ"] = y["tab_typ"];
                //        x["mdf_typ"] = j[0]["tab_typ"];
                //    }
                //    if (Convert.ToInt32(y["tab_len"]) != Convert.ToInt32(j[0]["tab_len"]))
                //    {
                //        b = true;
                //        x["csv_len"] = y["tab_len"];
                //        x["mdf_len"] = j[0]["tab_len"];
                //    }
                //    if (Convert.ToInt32(y["tab_lun"]) != Convert.ToInt32(j[0]["tab_lun"]))
                //    {
                //        b = true;
                //        x["csv_lun"] = Convert.ToDecimal(y["tab_lun"]);
                //        x["mdf_lun"] = Convert.ToDecimal(j[0]["tab_lun"]);
                //    }
                //    if (Convert.ToInt32(y["tab_dec"]) != Convert.ToInt32(j[0]["tab_dec"]))
                //    {
                //        b = true;
                //        x["csv_dec"] = Convert.ToDecimal(y["tab_dec"]);
                //        x["mdf_dec"] = Convert.ToDecimal(j[0]["tab_dec"]);
                //    }
                //    if (Convert.ToInt32(y["tab_idn"]) != Convert.ToInt32(j[0]["tab_idn"]))
                //    {
                //        b = true;
                //        x["csv_idn"] = Convert.ToDecimal(y["tab_idn"]);
                //        x["mdf_idn"] = Convert.ToDecimal(j[0]["tab_idn"]);
                //    }
                //    if (b)
                //        tDif.Rows.Add(x);
                //}
            }

            //foreach (DataRow y in tMdfTab.Rows)
            //{
            //    s = (string)y["tab_col"];

            //    j = tCsvTab.Select("tab_tab='" + (string)y["tab_tab"] + "' AND tab_col='" + (string)y["tab_col"] + "'");

            //    if (j.Length == 0)
            //    {
            //        x = tDif.NewRow();
            //        x["dif_tip"] = "TAB";
            //        x["dif_tab"] = y["tab_tab"];
            //        x["dif_col"] = y["tab_col"];
            //        x["csv_typ"] = y["tab_typ"];
            //        x["csv_len"] = y["tab_len"];
            //        x["csv_lun"] = y["tab_lun"];
            //        x["csv_dec"] = y["tab_dec"];
            //        x["csv_idn"] = y["tab_idn"];
            //        tDif.Rows.Add(x);
            //    }
            //}

            dgv5.DataSource = tDif;
        }

        private void button6_Click(object sender, EventArgs e)
        {
            string sTit = "Differenze";

            string sFil = "";

            string s = "";

            s += "dif_tip, Tipo, 20, StringLiteral;";
            s += "dif_tab, Tabella, 80, Integer;";
            s += "dif_col, Field, 50, StringLiteral;";
            s += "csv_typ, csv Tipo, 50, StringLiteral;";
            s += "mdf_typ, mdf Tipo, 50, StringLiteral;";
            s += "csv_len, csv Lung, 30, StringLiteral;";
            s += "mdf_len, mdf Lung, 30, StringLiteral;";
            s += "csv_lun, csv Lung, 30, StringLiteral;";
            s += "mdf_lun, mdf Lung, 30, StringLiteral;";
            s += "csv_dec, csv Decim, 30, StringLiteral;";
            s += "mdf_dec, mdf Decim, 30, StringLiteral;";
            s += "csv_idn, csv Itentity, 30, StringLiteral;";
            s += "mdf_idn, mdf Itentity, 30, StringLiteral;";

            for (int i = 0; i <= 100; i++)
            {
                sFil = "C:\\APproject\\PDF\\DiffMdf_" + i.ToString() + ".xls";
                if (File.Exists(sFil))
                {
                    try
                    {
                        File.Delete(sFil);
                        break;
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message); 
                    }

                }
                else
                    break;
            }

            DataSet ds = new DataSet();

            DataTable t = (DataTable)dgv5.DataSource;

            //foreach (DataColumn y in t.Columns)
            //{
            //    if (Convert.ToString(t.Rows[2][y.ColumnName]).Trim() == "-1" || Convert.ToString(t.Rows[2][y.ColumnName]).Trim() == "X")
            //        s += y.ColumnName + ", " + y.ColumnName + ", 50, StringLiteral;";
            //}

            DataTable tTit = (new clsExcel()).FillTitStr(s, "TabTit");

            ds.Tables.Add(tTit);
            ds.Tables.Add(t);

            (new clsExcel()).exportToExcel(ds, sFil, sTit, "", true);

            ds.Tables.Remove(t);
            ds.Clear();
            ds.Dispose();
        }

    }
}
