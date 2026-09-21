using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Diagnostics;

namespace APOffice
{
    class clsExcel2
    {
        clsDefine _clsDef = new clsDefine();
        clsFuncs _clsFun = new clsFuncs();

        private string _strPath = "";

        public clsExcel2()
        {
            string s = "C:\\APproject\\XLS";
            if (!Directory.Exists(s))
                Directory.CreateDirectory(s);
            _strPath = s + "\\";
        }

        private void Processo(string strCmd)
        {
            System.Diagnostics.Process pr = new Process();
            ProcessStartInfo ps = new ProcessStartInfo(strCmd);
            pr.StartInfo = ps;
            ps.UseShellExecute = true;
            pr.Start();
        }

        public void exportToXls1(DataTable tabTab, string strFld, string strFil, string strTit, string strFoo)
        {
            //string sCfi = "04269370286";

            string sFil = _strPath + DateTime.Now.ToString("yyyyMMddmmss") + "_" + strFil;
            if (File.Exists(_strPath + sFil))
                File.Delete(_strPath + sFil);
            File.Copy(_strPath + "Modelli\\" + strFil, sFil);

            string sXls = sFil; // "C:\\SouNet2010\\GesConty\\Xls\\ComPolivalente\\STANDARD COMUNICAZIONE POLIVALENTE Q_FE.XLS";
            //string sShe = "";

            System.Data.OleDb.OleDbConnection cn;
            OleDbCommand cm = new OleDbCommand();
            cn = new System.Data.OleDb.OleDbConnection("provider=Microsoft.Jet.OLEDB.4.0;Data Source='" + sXls + "';Extended Properties='Excel 8.0;HDR=No'");
            cm.Connection = cn;
            cn.Open();

            //cmd.Connection = con;
            //con.Open();

            //DataTable dtExcelSchema = cn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);

            //sShe = dtExcelSchema.Rows[0]["TABLE_NAME"].ToString();

            DataTable sch = cn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
            string sTab = (string)sch.Rows[0][2];

            string s = "SELECT * FROM [" + sTab + "]";  //Sheet1
            cm.CommandText = s;
            OleDbDataReader rs = cm.ExecuteReader();
            DataTable t = new DataTable();
            t.TableName = "TabXls";
            t.Load(rs);

            //dgv1.DataSource = t;

            //DataView v = new DataView(tabTab, "pno_tip='" + strTip + "'", "pno_cfd, pno_ddo, pno_ndo", DataViewRowState.CurrentRows);

            string[] aFld = strFld.Split(';');

            //Titolo

            s = "UPDATE [" + sTab + "] SET F1='" + strTit + "', ";

            //F4=null, F5=null, F31=null, F32=null 

            for (int i = 2; i < aFld.Length; i++)
                s += "F" + i.ToString() + "=null,";

            s = s.Substring(0, s.Length - 1);

            s += " WHERE F1='TITOLO'";
            try
            {
                cm.CommandText = s;
                cm.ExecuteNonQuery();

            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }

            int r = 0;

            //Corpo

            foreach (DataRow y in tabTab.Rows)
            {
                r++;

                if (r == 1)
                {
                    s = "UPDATE [" + sTab + "] SET ";
                    //s += "(F1, F2, F3, F4, F5, F7, F8, F9, F28, F29, F30, F31, F32) VALUES(";
                    //s += "(";

                    //for (int i = 1; i < aFld.Length; i++)
                    //{

                    int i = 0;

                    foreach (string ss in aFld)
                    {
                        string sVal = "";

                        if (ss != "")
                        {
                            i++;

                            string[] a = ss.Split(',');
                            if (a[3].Trim() == "StringLiteral")
                            {
                                if (DBNull.Value.Equals(y[a[0]]))
                                    sVal = "''";
                                else
                                    sVal = "'" + _clsFun.FaiLApice((string)y[a[0]]) + "'";

                                s += "F" + i.ToString() + "=" + sVal + ",";
                            }
                            else if (a[3].Trim() == "Decimal")
                            {
                                if (DBNull.Value.Equals(y[a[0]]))
                                    sVal = "0";
                                else
                                    sVal = Convert.ToDecimal(y[a[0]]).ToString("0.00").Replace(",", ".") + "";
                                //sVal = ((Decimal)y[a[0]]).ToString("0.00").Replace(",", ".") + "";

                                s += "F" + i.ToString() + "=" + sVal + ",";
                            }
                            else if (a[3].Trim() == "DateTime")
                            {
                                if (DBNull.Value.Equals(y[a[0]]))
                                    sVal = "''";
                                else
                                    sVal = "'" + ((DateTime)y[a[0]]).ToString("dd/MM/yy") + "'";

                                s += "F" + i.ToString() + "=" + sVal + ",";
                            }
                            else
                                Console.WriteLine("xxxxxxxxxxxxx");
                        }
                    }

                    //}

                    //s = s.Substring(0, s.Length - 1);
                    //s += ") VALUES(";

                    s = s.Substring(0, s.Length - 1);

                    //s += ")";

                    s += " WHERE F1='XYZ'";
                }
                else
                {
                    s = "INSERT INTO [" + sTab + "] ";
                    //s += "(F1, F2, F3, F4, F5, F7, F8, F9, F28, F29, F30, F31, F32) VALUES(";
                    s += "(";

                    for (int i = 1; i < aFld.Length; i++)
                        s += "F" + i.ToString() + ",";

                    s = s.Substring(0, s.Length - 1);

                    s += ") VALUES(";

                    foreach (string ss in aFld)
                    {
                        if (ss != "")
                        {
                            string[] a = ss.Split(',');
                            if (a[3].Trim() == "StringLiteral")
                            {
                                if (DBNull.Value.Equals(y[a[0]]))
                                    s += "'',";
                                else
                                    s += "'" + _clsFun.FaiLApice((string)y[a[0]]) + "',";
                            }
                            else if (a[3].Trim() == "Decimal")
                            {
                                if (DBNull.Value.Equals(y[a[0]]))
                                    s += "0,";
                                else
                                    s += Convert.ToDecimal(y[a[0]]).ToString("0.00").Replace(",", ".") + ",";
                            }
                            else if (a[3].Trim() == "DateTime")
                            {
                                if (DBNull.Value.Equals(y[a[0]]))
                                    s += "'',";
                                else
                                    s += "'" + ((DateTime)y[a[0]]).ToString("dd/MM/yy") + "',";
                            }
                            else
                                Console.WriteLine("xxxxxxxxxxxxx");
                        }
                    }

                    s = s.Substring(0, s.Length - 1);

                    s += ")";
                }

                try
                {
                    cm.CommandText = s;
                    cm.ExecuteNonQuery();

                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.Message);
                }

                //break;

            }

            //Piede
            if(strFoo != "")
            {
                s = "";
                string[] aa = strFoo.Split(';');

                s = "INSERT INTO [" + sTab + "] ";
                //s += "(F1, F2, F3, F4, F5, F7, F8, F9, F28, F29, F30, F31, F32) VALUES(";
                s += "(";

                for (int i = 1; i < aFld.Length; i++)
                    s += "F" + i.ToString() + ",";

                s = s.Substring(0, s.Length - 1);

                s += ") VALUES(";

                //for (int i = 1; i < aFld.Length; i++)
                //{
                //    Boolean b = true;
                //    foreach(string ss in aa)
                //    {
                //        string[] aaa = ss.Split(',');

                //        if(aaa[0] == i.ToString())
                //        {
                //            b = false;
                //            s += "" + aaa[1] + ",";
                //        }
                //    }

                //    if(b)
                //        s += "null,";
                //}

                string[] aaa = strFoo.Split(',');

                Console.WriteLine("xxx");


                //s = "";
                foreach (string ss in aaa)
                {
                    Console.WriteLine(ss);
                    if (ss == "")
                        s += "null,";
                    else
                        s += ss + ",";
                }

                s = s.Substring(0, s.Length - 1);
                s += ")";

                try
                {
                    cm.CommandText = s;
                    cm.ExecuteNonQuery();

                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.Message);
                }

            }

            cn.Close();

            Processo(sFil);
        }

    }
}

