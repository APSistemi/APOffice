using System;
using System.Collections.Generic;
using System.Collections;
using System.Linq;
using System.Text;
using System.Data;
using System.Data.OleDb;
using System.Windows.Forms;
using System.IO;

namespace APOffice
{
    class clsDivSMA
    {
        clsDefine _clsDef = new clsDefine();
        clsFuncs _clsFun = new clsFuncs();

        public clsDivSMA()
        {}

        public DataRow DivFat(string strRig, DataRow rowArt, DataTable tabLis, ref string strMsg)
        {
            string s = "";
            string sMsg = "";

            String sArf = strRig.Substring(25, 9);
            string sRep = strRig.Substring(22, 3);
            string sQta = strRig.Substring(49, 11);
            string sNdo = strRig.Substring(99, 6);
            string sDdo = strRig.Substring(105, 10);
            string sPrv = strRig.Substring(131, 14);
            string sPrc = strRig.Substring(145, 14);

            string sIva = "";
            string sEan = "";
            string sDes = "";
            string sTgr = "";
            string sPne = "";
            string sUmi = "";
            string sPxc = "";

            rowArt["div_dva"] = DateTime.Today.ToShortDateString();  //OK
            rowArt["div_arf"] = sArf;
            rowArt["div_cos"] = _clsFun.Txt2Dec(sPrc);
            rowArt["div_prp"] = _clsFun.Txt2Dec(sPrv);
            rowArt["for_rep"] = sRep;
            rowArt["div_tpd"] = "F";

            DataRow[] j = tabLis.Select("sma_arf='" + sArf + "'");
            if (j.Length == 0)
            {
                strMsg += "Non trovato " + sArf + " qta " + sQta + _clsDef.CRLF;
            }
            else
            {
                if (sArf == "000182490")
                    Console.WriteLine("zzzzzzzzz");

                sEan = (string)j[0]["sma_ean"];
                sDes = (string)j[0]["sma_ard"];
                sTgr = (string)j[0]["sma_tgr"];
                sIva = (string)j[0]["sma_iva"];
                decimal dPne = Convert.ToDecimal(j[0]["sma_pne"]);
                sUmi = (string)j[0]["sma_umi"];
                //sPxc = (string)j[0][10];
                decimal dPxc = Convert.ToDecimal(j[0]["sma_pxc"]);
                //if (!DBNull.Value.Equals(j[0][9]))
                //    dPxc = Convert.ToDecimal(j[0][9]);

                rowArt["div_ard"] = sDes;

                rowArt["for_tgr"] = sTgr;

                rowArt["div_iva"] = sIva;

                rowArt["for_ecr"] = "";

                rowArt["for_iva"] = sIva.PadLeft(3, Convert.ToChar('0'));

                rowArt["for_umi"] = sUmi;

                rowArt["div_pxc"] = dPxc;

                rowArt["div_pne"] = dPne;

                rowArt["div_qta"] = _clsFun.Txt2Dec(sQta);

                rowArt["div_cof"] = 0;

                rowArt["div_imp"] = 0;

                rowArt["div_tva"] = "";

                rowArt["div_ean"] = sEan;

                if (sEan == "")
                    strMsg += "Barcode non definito in " + sArf + " qta " + sQta + _clsDef.CRLF;

            }

            return rowArt;
        }

        public void DivSmaAna(string strRig, DataTable tabSma)
        {
            string s = "";
            DataRow[] j;

            ArrayList aWhe = new ArrayList();
            aWhe.Add("sma_arf");
            ArrayList aExl = new ArrayList();

            string sCon = _clsFun.ConMdb(Path.GetDirectoryName(_clsDef.FILEINI) + "\\ApSMA.mdb");

            if (strRig.Length > 100 && strRig.Substring(0, 2) == "01")
            {
                String sArf = strRig.Substring(16, 9);
                string sArd = strRig.Substring(25, 30).Trim();
                string sRep = strRig.Substring(68, 3);
                string sTgr = strRig.Substring(84, 2);
                string sUmi = strRig.Substring(86, 2);

                if(sArf == "000262616")
                    Console.WriteLine("ZZZ");

                //decimal dPxc = 0;
                //s = strRig.Substring(89, 4).Trim();
                //if (_clsFun.Numerico(s, "0123456789.,"))
                //    dPxc = Convert.ToDecimal(s);

                DataRow x = tabSma.NewRow();
                x["sma_arf"] = sArf;
                x["sma_ard"] = sArd;
                //x["sma_rep"] = sRep;
                x["sma_umi"] = sUmi;
                x["sma_tgr"] = sTgr;
                //x["sma_pxc"] = dPxc;
                //x["sma_ean"] = (string)y[12];

                if (sUmi != "KG")
                    sUmi = "NR";

                j = tabSma.Select("sma_arf='" + x["sma_arf"] + "'");
                if (j.Length > 0)
                {
                    //s = _clsFun.SqlUpdRow("AnaSma", tabSma, j[0], x, aWhe, aExl);

                    s = "";
                    Boolean b = false;
                    if (((string)j[0]["sma_ard"]).Trim() != sArd)
                        b = true;
                    if ((string)j[0]["sma_umi"] != sUmi)
                        b = true;
                    if ((string)j[0]["sma_tgr"] != sTgr)
                        b = true;

                    if (b)
                    {
                        s = "UPDATE AnaSma SET ";
                        s += "sma_ard='" + _clsFun.FaiLApice(sArd) + "', ";
                        s += "sma_umi='" + sUmi + "', ";
                        s += "sma_tgr='" + sTgr + "' ";
                        s += "WHERE sma_arf='" + sArf + "'";
                    }
                }
                else
                {
                    s = _clsFun.SqlInsertRow("AnaSma", tabSma, x);
                    tabSma.Rows.Add(x);
                }
                if (s != "")
                    _clsFun.MdbWrite(s, sCon);
            }
            else if (strRig.Length > 40 && strRig.Substring(0, 2) == "04")
            {
                String sArf = strRig.Substring(16, 9);
                string sEan = strRig.Substring(26, 13);

                j = tabSma.Select("sma_arf='" + sArf + "'");
                if (j.Length > 0)
                {
                    if (DBNull.Value.Equals(j[0]["sma_ean"]) || (string)j[0]["sma_ean"] != sEan)
                    {
                        s = "UPDATE AnaSma SET sma_ean='" + sEan + "' WHERE sma_arf='" + sArf + "'";
                        _clsFun.MdbWrite(s, sCon);
                    }
                }

                Console.WriteLine("ZZZ");
            }
            else if (strRig.Length > 40 && strRig.Substring(0, 2) == "05")
            {
                String sArf = strRig.Substring(16, 9);
                string sIva = strRig.Substring(36).Trim();

                if (_clsFun.Numerico(sIva, "0123456789."))
                {
                    string[] a = sIva.Split('.');
                    sIva = a[0];
                    sIva = sIva.PadLeft(3,Convert.ToChar("0"));

                    j = tabSma.Select("sma_arf='" + sArf + "'");
                    if (j.Length > 0)
                    {
                        if (DBNull.Value.Equals(j[0]["sma_iva"]) || (string)j[0]["sma_iva"] != sIva)
                        {
                            s = "UPDATE AnaSma SET sma_iva='" + sIva + "' WHERE sma_arf='" + sArf + "'";
                            _clsFun.MdbWrite(s, sCon);
                        }
                    }
                }
                Console.WriteLine("ZZZ");
            }
        }

        public DataTable Xls2Tab(string strTip, string strFil, ProgressBar progressBar1)
        {
            //s = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + s + ";Extended Properties=\"Excel 8.0;HDR={1};IMEX=1\"";

            string sCon = _clsFun.ConMdb(Path.GetDirectoryName(_clsDef.FILEINI) + "\\ApSMA.mdb");
            string s = "SELECT * FROM AnaSma ORDER BY sma_arf";
            DataTable tSma = _clsFun.FillTabMdb("AnaSma", s, false, sCon);
            DataColumn[] keys = new DataColumn[1];
            keys[0] = tSma.Columns["sma_arf"];
            tSma.PrimaryKey = keys;

            if (File.Exists(strFil))
            {
                DataRow[] j;

                //s = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + s + ";Extended Properties=Excel 8.0;";
                s = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + strFil + ";Extended Properties=\"Excel 8.0;HDR={1};IMEX=1\"";

                OleDbConnection cn = new System.Data.OleDb.OleDbConnection(s);
                cn.Open();

                //DataTable sch = cn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, new object[] { null, null, null, "TABLE" });
                DataTable sch = cn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                string sTab = (string)sch.Rows[0][2];

                OleDbCommand cm = new OleDbCommand();
                cm.Connection = cn;
                //s = "SELECT * FROM [" + sTab + "] WHERE Referenza <> null AND EAN <> null";
                //s = "SELECT * FROM [" + sTab + "] WHERE Cod. referenza <> null AND Bar code <> null";
                s = "SELECT * FROM [" + sTab + "] WHERE 'Cod# referenza' <> null AND 'Bar code' <> null";

                cm.CommandText = s;
                OleDbDataReader dr = cm.ExecuteReader();
                DataTable tXls = new DataTable();
                tXls.TableName = "TabXls";
                tXls.Load(dr);

                keys = new DataColumn[2];
                keys[0] = tXls.Columns["Cod# referenza"];
                keys[1] = tXls.Columns["Bar code"];
                tXls.PrimaryKey = keys;

                //DataTable t = new clsGenTabTmp().TabTmpDivFor("TabSma");
                DataRow x;
                //string sTgr = "";
                //decimal dPne = 0;

                progressBar1.Value = 0;
                progressBar1.Maximum = tXls.Rows.Count;
                progressBar1.Minimum = 0;

                ArrayList aWhe = new ArrayList();
                aWhe.Add("sma_arf");
                ArrayList aExl = new ArrayList();

                foreach (DataRow y in tXls.Rows)
                {
                    progressBar1.Increment(1);
                    System.Windows.Forms.Application.DoEvents();

                    x = tSma.NewRow();
                    x["sma_arf"] = (string)y[10];
                    x["sma_ard"] = (string)y[11];
                    x["sma_ean"] = (string)y[12];

                    x["sma_iva"] = "";
                    s = (string)y[18];
                    if (_clsFun.Numerico(s, "0123456789,."))
                        s = Convert.ToDecimal(s).ToString("000");
                    x["sma_iva"] = s;

                    s = (string)y[19];

                    Console.WriteLine("xxx");

                    x["sma_pne"] = 0;
                    x["sma_tgr"] = "KG";

                    string[] a = s.Split('/');
                    if (a.Length > 1)
                    {
                        s = a[0].Trim();
                        Console.WriteLine("xxx");

                        if (_clsFun.Numerico(s, "0123456789.,"))
                            x["sma_pne"] = Convert.ToDecimal(s); // .Replace(",",".");
                        x["sma_tgr"] = a[1].Trim();
                    }

                    x["sma_umi"] = "NR";

                    x["sma_pxc"] = 1;
                    s = Convert.ToString(y[22]);
                    if (_clsFun.Numerico(s, "0123456789.,"))
                        x["sma_pxc"] = Convert.ToString(s);

                    j = tSma.Select("sma_arf='" + x["sma_arf"] + "'");
                    if (j.Length > 0)
                    {
                        s = _clsFun.SqlUpdRow("AnaSma", tSma, j[0], x, aWhe, aExl);
                    }
                    else
                    {
                        s = _clsFun.SqlInsertRow("AnaSma", tSma, x);
                        tSma.Rows.Add(x);
                    }
                    if (s != "")
                        _clsFun.MdbWrite(s, sCon);

                }

                tXls.Dispose();
                cm.Cancel();

                s = strFil + "_" + DateTime.Now.ToString("yyyyMMddHHmmss");
                cm.Dispose();
                cn.Close();
                cn.Dispose();

                File.Move(strFil, s);
            }

            return tSma;

        }

    }
}
