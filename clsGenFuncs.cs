using System;
using System.Collections;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using System.Data.SqlClient;
using System.Data.OleDb;
using System.Windows.Forms;
using System.IO;

namespace APOffice
{
    class clsFuncs
    {
        private clsDefine _clsDef = new clsDefine();

        public string ConSql(string strSql)
        {
            string s = strSql;
            if (Numerico(s) && s == "5")
            {
                s = FileIni("R", clsDefine.enuIni.IniAPOfficeSql, "");
                if(s.Contains("ApOffice"))
                    s = s.Replace("ApOffice", "ApStatStor");
                if (s.Contains("APOffice"))
                    s = s.Replace("APOffice", "ApStatStor");
            }
            else if (Numerico(s) && s == "4")
            {
                s = FileIni("R", clsDefine.enuIni.IniAPOfficeSql, "");
                s = s.Replace("APOffice", "ApLog");
            }
            else if (Numerico(s) && s == "3")
                s = FileIni("R", clsDefine.enuIni.IniAPOfficeSqlStat, "");
            else if (s == "")
                s = FileIni("R", clsDefine.enuIni.IniAPOfficeSql, "");

            ErrorLog("INI SQL", s);

            return s;
        }

        public string ConMdb(string strMdb)
        {
            string s = "";
            if (strMdb == "")
            {
                strMdb = _clsDef.MDBGEN;
                s = FileIni("R", clsDefine.enuIni.IniAPOfficeMdb, "");
            }
            s = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source='" + s + strMdb + "';";
            return s;
        }

        public string FileIni(string strAzi, clsDefine.enuIni iniCod, string strPar)
        {
            string sCod = Convert.ToInt16(iniCod).ToString("000");
            string sPar = "";
            string sRig = "";

            if (!File.Exists(_clsDef.FILEINI))
                ErrorLog("INI non trovato", _clsDef.FILEINI);

            if ((strAzi == "R" || strAzi == "W") && File.Exists(_clsDef.FILEINI))
            {
                using (StreamReader sr = new StreamReader(_clsDef.FILEINI))
                {
                    while ((sRig = sr.ReadLine()) != null)
                    {
                        if (sRig.Length > 3 && sRig.Substring(0, 3) == sCod)
                        {
                            if (sRig.Length > 4)
                                sPar = sRig.Substring(4).Trim();
                            break;
                        }
                    }
                    sr.Close();
                    sr.Dispose();
                }
            }
            if (strAzi == "W" && sPar != strPar)
            {
                bool b = true;
                string path = Path.GetDirectoryName(_clsDef.FILEINI);
                if (!Directory.Exists(path))
                {
                    Directory.CreateDirectory(path);
                    if (!Directory.Exists(path))
                    {
                        int num = (int)MessageBox.Show("Impossibile creare la directory " + path);
                        b = false;
                    }
                }
                if (b)
                {
                    StreamWriter sw = new StreamWriter(_clsDef.FILEINI + "tmp");
                    if (File.Exists(_clsDef.FILEINI))
                    {
                        using (StreamReader sr = new StreamReader(this._clsDef.FILEINI))
                        {
                            b = false;
                            while ((sRig = sr.ReadLine()) != null)
                            {
                                if (sRig.Substring(0, 3) == sCod)
                                {
                                    sRig = sCod + " " + strPar;
                                    b = true;
                                }
                                sw.Write(sRig + "\r\n");
                            }
                            if (!b)
                            {
                                sRig = sCod + " " + strPar;
                                sw.Write(sRig + "\r\n");
                            }
                            sr.Close();
                            sr.Dispose();
                        }
                    }
                    else
                    {
                        sRig = sCod + " " + strPar;
                        sw.Write(sRig + "\r\n");
                    }
                    ((TextWriter)sw).Flush();
                    sw.Close();
                    sw.Dispose();
                    File.Delete(_clsDef.FILEINI);
                    File.Move(_clsDef.FILEINI + "tmp", _clsDef.FILEINI);
                }
            }

            ErrorLog("INI SQL 3", sPar);

            return sPar;
        }

        public string FileIniIni(string strCod)
        {
            string sRes = "";

            if (File.Exists("Ini.ini"))
            {
                using (StreamReader sr = new StreamReader("Ini.ini"))
                {
                    string sRig = "";

                    while ((sRig = sr.ReadLine()) != null)
                    {
                        if (sRig.Length > 3 && sRig.Substring(0, 3) == strCod)
                        {
                            if (sRig.Length > 4)
                                sRes = sRig.Substring(4).Trim();
                            break;
                        }
                    }
                    sr.Close();
                    sr.Dispose();
                }
            }

            return sRes;
        }

        public DataRow RowIniz(DataTable tTab)
        {
            DataRow dataRow = tTab.NewRow();
            foreach (DataColumn dataColumn in (InternalDataCollectionBase)tTab.Columns)
            {
                Console.WriteLine((object)dataColumn.DataType);
                if (((object)dataColumn.DataType.Name).ToString() == "String")
                {
                    if (DBNull.Value.Equals(dataRow[dataColumn.ColumnName]))
                        dataRow[dataColumn.ColumnName] = (object)"";
                }
                else if (dataColumn.DataType.Name == "Boolean")
                {
                    if (DBNull.Value.Equals(dataRow[dataColumn.ColumnName]))
                        dataRow[dataColumn.ColumnName] = (object)false;
                }
                else if (dataColumn.DataType.Name == "Decimal")
                {
                    if (DBNull.Value.Equals(dataRow[dataColumn.ColumnName]))
                        dataRow[dataColumn.ColumnName] = (object)0;
                }
                else if (dataColumn.DataType.Name.Substring(0, 3) == "Int")
                {
                    if (DBNull.Value.Equals(dataRow[dataColumn.ColumnName]))
                        dataRow[dataColumn.ColumnName] = (object)0;
                }
                else if (dataColumn.DataType.Name == "DateTime")
                {
                    if (DBNull.Value.Equals(dataRow[dataColumn.ColumnName]))
                        dataRow[dataColumn.ColumnName] = (object)this._clsDef.DAYOUT;
                }
                else if (dataColumn.DataType.Name != "Guid")
                {
                    int num = (int)MessageBox.Show("c.DataType.Name " + dataColumn.DataType.Name + " non definito su " + tTab.TableName);
                }
            }
            return dataRow;
        }

        public DataRow TypeDefault(DataTable tab, DataRow row)
        {
            foreach (DataColumn c in tab.Columns)
            {
                if (DBNull.Value.Equals(row[c]))
                {
                    Console.WriteLine((object)c.DataType);
                    if (((object)c.DataType.Name).ToString() == "String")
                        row[c.ColumnName] = (object)"";
                    else if (c.DataType.Name == "Boolean")
                        row[c.ColumnName] = (object)false;
                    else if (c.DataType.Name == "Decimal")
                        row[c.ColumnName] = (object)0;
                    else if (c.DataType.Name.Substring(0, 3) == "Int")
                        row[c.ColumnName] = (object)0;
                    else if (c.DataType.Name == "DateTime")
                        row[c.ColumnName] = (object)this._clsDef.DAYOUT;
                    else if (c.DataType.Name != "Guid")
                        MessageBox.Show("c.DataType.Name " + c.DataType.Name + " non definito su " + tab.TableName);
                }
            }
            return row;
        }

        public string NewNum(string strYea, clsDefine.enuNumeratori enuCod, int intLen, string strCon)
        {
            string sCod = Convert.ToInt16((object)enuCod).ToString("000");
            string sNum = NumGet(strYea, enuCod, ConSql(""));
            if (sNum == "")
                sNum = "0";
            sNum = Convert.ToString(Convert.ToInt32(sNum) + 1);
            NumSet(strYea, enuCod, sNum, strCon);
            return sNum.PadLeft(intLen, Convert.ToChar("0"));
        }

        public string NumGet(string strYea, clsDefine.enuNumeratori enuCod, string strCon)
        {
            string sCod = Convert.ToInt16((object)enuCod).ToString("000");
            string p = "TabNumeratori";
            string s = "SELECT * FROM " + p + " WHERE tab_yea ='" + strYea + "' AND tab_cod ='" + sCod + "'";
            DataTable t = FillTabSql(p, s, true, strCon);
            s = "";
            if (t.Rows.Count > 0)
                s = Convert.ToString(t.Rows[0]["tab_val"]).Trim();
            else 
            {
                MessageBox.Show("Aggiornamento numeratore " + strYea + "-" + sCod + "", "CONTROLLO NUMERATORI");

                string sDes = "NUMERATORE";

                if (sCod == "004")
                    sDes += " OFFERTE";
                else if (sCod == "005")
                    sDes += " MOVIMENTI";
                else if (sCod == "006")
                    sDes += " MOV FATTURE";
                else if (sCod == "007")
                    sDes += " DOC FATTURE ";
                else if (sCod == "008")
                    sDes += " DDT";
                else if (sCod == "009")
                    sDes += " ORD FORNITORI";
                else if (sCod == "011")
                    sDes += " LOTTI";
                else if (sCod == "012")
                    sDes += " TESSERE";
                else if (sCod == "013")
                    sDes += " DOC MOVIMENTI 2";
                else if (sCod == "014")
                    sDes += " FAT ELETTRONICHE";

                s = "SELECT * FROM " + p + " WHERE tab_cod='" + sCod + "'";
                t = FillTabSql(p, s, true, strCon);

                DataRow y = t.NewRow();
                y["tab_yea"] = strYea;
                y["tab_cod"] = sCod;
                y["tab_des"] = sDes;
                y["tab_val"] = 0;

                s = SqlInsertRow(p, t, y);
                SqlWrite(s, strCon);
                s = "0";
            }
            return s;
        }

        public void NumSet(string strYea, clsDefine.enuNumeratori enuCod, string strVal, string strCon)
        {
            string sCod = Convert.ToInt16((object)enuCod).ToString("000");
            string p = "TabNumeratori";
            string s = "UPDATE TabNumeratori SET tab_val=" + strVal + " WHERE tab_yea ='" + strYea + "' AND tab_cod ='" + sCod + "'";
            SqlWrite(s, strCon);
        }

        public string ParGet(clsDefine.enuParametri epaCod, string strCon)
        {
            string sCod = Convert.ToInt16((object)epaCod).ToString("000");
            string sTab = "TabParametri";
            string s = "SELECT * FROM " + sTab + " WHERE tab_cod ='" + sCod + "'";
            DataTable t = FillTabSql(sTab, s, true, strCon);
            s = "";
            if (t.Rows.Count > 0)
                s = Convert.ToString(t.Rows[0]["tab_val"]);
            return s;
        }

        public void ParSet(clsDefine.enuParametri epaCod, string strPar, string strCon)
        {
            string sTab = "TabParametri";
            string sCod = Convert.ToString(epaCod).PadLeft(3, Convert.ToChar("0"));
            string s = "SELECT * FROM " + sTab + " WHERE tab_cod ='" + sCod + "'";
            DataTable t = FillTabSql(sTab, s, true, strCon);
            if (t.Rows.Count == 0)
            {
                if (MessageBox.Show("Parametro " + sCod + " non presente, inserisco?", "GESTIONE PARAMETRI", MessageBoxButtons.YesNo, MessageBoxIcon.Question) != DialogResult.Yes)
                    return;
                DataRow x = t.NewRow();
                x["tab_cod"] = sCod;
                x["tab_des"] = (object)("Parametro " + sCod);
                x["tab_val"] = (object)strPar;
                t.Rows.Add(x);
                s = SqlInsertRow(sTab, t, x);
                if (s != "")
                    SqlWrite(s, strCon);
            }
            else
                SqlWrite("UPDATE " + sTab + " SET tab_val='" + strPar + "' WHERE tab_cod ='" + sCod + "'", strCon);
        }

        public Boolean TestMdf(string sDcn)
        {
            Boolean b = true;
            try
            {
                SqlConnection cn = new SqlConnection(sDcn);
                cn.Open();
                cn.Dispose();
            }
            catch (Exception ex)
            {
                b = false;
                ErrorLog(ex.Message, sDcn);
                //this._clsMail.InvMail("edp@spacrovigo.org", "EDP", "paolo.secchettin@spacrovigo.org", "", "Errori su SpacServer", "clsGenFun.SqlWrite", "", "");
            }

            return b;
        }

        public DataSet FillDataSetSql(string sSql, string sDcn)
        {
            DataSet ds = new DataSet();
            try
            {
                if (sDcn == "") sDcn = this.ConSql("");
                using (SqlConnection cn = new SqlConnection(sDcn))
                {
                    SqlDataAdapter da = new SqlDataAdapter(sSql, cn);
                    da.Fill(ds);
                }
            }
            catch (Exception ex)
            {
                this.ErrorLog(ex.Message, sSql);
            }
            return ds;
        }

        public DataTable FillTabSql(string sTab, string sSql, bool bPrimo, string sDcn)
        {
            //FileLog("ApBellia", "Passo 2", sDcn);

            DataTable t = new DataTable(sTab);
            try
            {
                if (sDcn == "")
                    sDcn = this.ConSql("");
                SqlConnection cn = new SqlConnection(sDcn);

                SqlCommand cm = new SqlCommand(sSql, cn);
                cm.CommandTimeout = 50;
                cn.Open();

                SqlDataReader dr = cm.ExecuteReader();
                DataTable sch = dr.GetSchemaTable();
                foreach (DataRow dataRow in (InternalDataCollectionBase)sch.Rows)
                {
                    DataColumn c = new DataColumn();
                    c.DataType = System.Type.GetType(Convert.ToString(dataRow[12]));
                    c.ColumnName = Convert.ToString(dataRow[0]);
                    c.AutoIncrement = false;
                    c.Caption = Convert.ToString(dataRow[0]);
                    c.ReadOnly = false;
                    c.Unique = false;

                    if (c.ColumnName == "art_pne")
                        Console.WriteLine("aaaa");
                    if (c.DataType.FullName == "System.String")
                        c.MaxLength = (int)Convert.ToInt32(dataRow[2]);
                    t.Columns.Add(c);
                }
                if (dr.HasRows)
                {
                    while (dr.Read())
                    {
                        DataRow row = t.NewRow();
                        for (int index = 0; index <= t.Columns.Count - 1; ++index)
                        {
                            if (DBNull.Value.Equals(dr[index]))
                            {
                                switch (t.Columns[index].DataType.Name)
                                {
                                    case "Decimal":
                                        row[index] = (object)0;
                                        continue;
                                    case "Int32":
                                        row[index] = (object)0;
                                        continue;
                                    case "String":
                                        row[index] = (object)"";
                                        continue;
                                    case "Boolean":
                                        row[index] = (object)false;
                                        continue;
                                    case "DateTime":
                                        row[index] = (object)new DateTime(2050, 1, 1);
                                        continue;
                                    case "Byte":
                                        row[index] = (object)0;
                                        continue;
                                    default:
                                        int num = (int)MessageBox.Show("Tipo dato non definito in RecTabRec su " + sTab + "!", "");
                                        continue;
                                }
                            }
                            else
                            {
                                if (t.Columns[index].DataType.Name == "String")
                                    row[index] = Convert.ToString(dr[index]).Trim();
                                else
                                    row[index] = dr[index];
                            }

                        }
                        t.Rows.Add(row);
                        if (bPrimo)
                            break;
                    }
                }
                cm.Dispose();
                dr.Close();
                cn.Dispose();

            }
            catch (Exception ex)
            {
                this.ErrorLog(ex.Message, sSql);
                //this._clsMail.InvMail("edp@spacrovigo.org", "EDP", "paolo.secchettin@spacrovigo.org", "", "Errori su SpacServer", "clsGenFun.SqlWrite", "", "");
            }

            return t;
        }

        public string SqlInsertRow(string sTab, DataTable tTmp, DataRow rTmp)
        {
            string str1 = "";
            string str2 = "";
            foreach (DataColumn dataColumn in (InternalDataCollectionBase)tTmp.Columns)
            {
                string colName = dataColumn.ColumnName;
                if (colName.EndsWith("Mdy") || colName == "LivDes" || colName == "LivTip" || colName == "LivMav" || colName == "LivMap" || colName == "CosFod" || colName == "EanMdy" || colName == "LiaMdy")
                    continue;

                if (colName.Length > 6 && colName.Substring(3, 4) == "_idx")
                    continue;
                if (colName.EndsWith("_idx"))
                    continue;

                str1 = str1 + colName + ",";
                str2 = str2 + FldDefault(rTmp, colName, true, false) + ",";
            }
            if (str1.Length == 0) return "";
            return "INSERT INTO " + sTab + " (" + str1.Substring(0, str1.Length - 1) + ") VALUES (" + str2.Substring(0, str2.Length - 1) + ")";
        }

        public string SqlUpdRow(string sTab, DataTable tTmp, DataRow rOri, DataRow rTmp, ArrayList aWhe, ArrayList aExl)
        {
            bool b = false;
            string sSql = "";
            string sWhe = "";
            foreach (DataColumn c in (InternalDataCollectionBase)tTmp.Columns)
            {
                string colName = c.ColumnName;
                if (colName.EndsWith("Mdy") || colName == "LivDes" || colName == "LivTip" || colName == "LivMav" || colName == "LivMap" || colName == "CosFod" || colName == "EanMdy" || colName == "LiaMdy")
                    continue;

                if (aWhe.IndexOf((object)colName) >= 0)
                    sWhe += colName + "=" + Convert.ToString(FldDefault(rTmp, colName, true, false)) + " AND ";
                else if (colName.Length > 6 && colName.Substring(3, 4) != "_idx") 
                {
                     //&& !FldDefault(rOri, c.ColumnName, false, false).Equals(FldDefault(rTmp, c.ColumnName, false, false)))
                    object o1 = FldDefault(rOri, colName, false, false);
                    object o2 = FldDefault(rTmp, colName, false, false);

                    if(!o1.Equals(o2))
                    {
                        sSql += colName + "=";
                        sSql += this.FldDefault(rTmp, colName, true, false);
                        sSql += ",";
                        if (aExl != null && aExl.IndexOf((object)colName) < 0)
                            b = true;
                    }
                }
            }
            if (b)
            {
                sSql = sSql.Substring(0, sSql.Length - 1);
                sWhe = sWhe.Substring(0, sWhe.Length - 4);
                sSql = "UPDATE " + sTab + " SET " + sSql + " WHERE " + sWhe;
            }
            else
                sSql = "";

            return sSql;
        }

        public string SqlUpdRowIdx(string sTab, DataTable tTmp, DataRow rOri, DataRow rTmp, ArrayList aExl)
        {
            bool b = false;
            string sSql = "";
            string sWhe = "";

            foreach (DataColumn c in tTmp.Columns)
            {
                string colName = c.ColumnName;
                if (colName.EndsWith("Mdy") || colName == "LivDes" || colName == "LivTip" || colName == "LivMav" || colName == "LivMap" || colName == "CosFod" || colName == "EanMdy" || colName == "LiaMdy")
                    continue;

                if (colName.Length > 6 && colName.Substring(3, 4) == "_idx")
                {
                    //if (DBNull.Value.Equals(rOri[c.ColumnName]))
                    //{
                    //    Console.WriteLine("xxxx");
                    //    break;
                    //}
                    //else
                    sWhe += (object)colName + "=" + ((int)rOri[colName]).ToString() + " AND ";
                }
                else if (colName.Length > 6 && colName.Substring(3, 4) != "_idx")
                {
                    //&& !FldDefault(rOri, c.ColumnName, false, false).Equals(FldDefault(rTmp, c.ColumnName, false, false)))
                    object o1 = FldDefault(rOri, colName, false, false);
                    object o2 = FldDefault(rTmp, colName, false, false);

                    if (!o1.Equals(o2))
                    {
                        sSql += colName + "=";
                        sSql += this.FldDefault(rTmp, colName, true, false);
                        sSql += ",";
                        if (aExl == null || aExl.IndexOf((object)colName) < 0)
                            b = true;
                    }
                }
            }
            if (b)
            {
                sSql = sSql.Substring(0, sSql.Length - 1);
                sWhe = sWhe.Substring(0, sWhe.Length - 4);
                sSql = "UPDATE " + sTab + " SET " + sSql + " WHERE " + sWhe;
            }
            else
                sSql = "";

            return sSql;
        }

        public object FldDefault(DataRow y, string x, bool bSql, bool bMdb)
        {
            int index = y.Table.Columns.IndexOf(x);
            try
            {
                if (y.Table.Columns[index].DataType.Name == "String")
                {
                    string strStr = "";
                    if (!DBNull.Value.Equals(y[index]))
                        strStr = Convert.ToString(y[index]).Trim();
                    int mLen = y.Table.Columns[index].MaxLength;
                    if (mLen > 0 && strStr.Length > mLen)
                        strStr = strStr.Substring(0, mLen);
                    if (bSql)
                        strStr = "'" + FaiLApice(strStr) + "'";
                    return (object)strStr.Trim();
                }
                else if (y.Table.Columns[index].DataType.Name == "DateTime")
                {
                    DateTime d = this._clsDef.DAYOUT;
                    if (!DBNull.Value.Equals(y[index]))
                        d = (DateTime)y[index];
                    if (!bSql)
                        return (object)d;
                    if (bMdb)
                        return (object)this.DayMdb(d);
                    else
                        return (object)this.DaySql(d);
                }
                else if (y.Table.Columns[index].DataType.Name == "Boolean")
                {
                    bool flag = false;
                    if (!DBNull.Value.Equals(y[index]))
                        flag = Convert.ToBoolean(y[index]);
                    if (!bSql)
                        return (object)flag;
                    if (flag)
                        return (object)"1";
                    else
                        return (object)"0";
                }
                else if (y.Table.Columns[index].DataType.Name == "Int32")
                {
                    Decimal num = 0;
                    if (!DBNull.Value.Equals(y[index]))
                        num = Convert.ToDecimal(y[index]);
                    //if (!bSql)
                    //    return (object)flag;
                    //if (flag)
                    //    return (object)"1";
                    //else
                    //    return (object)"0";
                    return num;
                }
                else
                {
                    if (!(y.Table.Columns[index].DataType.Name == "Decimal"))
                        return y[index];
                    Decimal num = new Decimal(0);
                    if (!DBNull.Value.Equals(y[index]))
                        num = Convert.ToDecimal(y[index]);
                    if (bSql)
                        return (object)Convert.ToString(num).Replace(',', '.');
                    else
                        return (object)num;
                }
            }
            catch (Exception ex)
            {
                ErrorLog(ex.Message, x);
                MessageBox.Show(ex.Message,x);
                return "";
            }

        }

        public DataTable FillTabMdb(string sTab, string sSql, bool bPrimo, string sDcn)
        {
            if (sDcn == "")
                sDcn = ConMdb("");
            OleDbConnection cn = new OleDbConnection(sDcn);
            OleDbCommand cm = new OleDbCommand(sSql, cn);
            DataTable t = new DataTable(sTab);
            try
            {
                cn.Open();
                OleDbDataReader oleDbDataReader = cm.ExecuteReader();
                DataTable schemaTable = oleDbDataReader.GetSchemaTable();
                foreach (DataRow dataRow in (InternalDataCollectionBase)schemaTable.Rows)
                {
                    DataColumn c = new DataColumn();
                    c.DataType = System.Type.GetType(Convert.ToString(dataRow[5]));
                    c.ColumnName = Convert.ToString(dataRow[0]);
                    c.AutoIncrement = false;
                    c.Caption = Convert.ToString(dataRow[0]);
                    c.ReadOnly = false;
                    c.Unique = false;
                    if (c.DataType.FullName == "System.String" && Convert.ToDouble(dataRow[2]) < 255)
                        c.MaxLength = (int)Convert.ToInt16(dataRow[2]);
                    t.Columns.Add(c);
                }
                if (oleDbDataReader.HasRows)
                {
                    while (oleDbDataReader.Read())
                    {
                        DataRow row = t.NewRow();
                        for (int index = 0; index <= schemaTable.Rows.Count - 1; ++index)
                        {
                            if (oleDbDataReader[index] == null)
                            {
                                switch (System.Type.GetTypeCode(row.GetType()))
                                {
                                    case TypeCode.Boolean:
                                        row[index] = (object)false;
                                        continue;
                                    case TypeCode.Decimal:
                                        row[index] = (object)0;
                                        continue;
                                    case TypeCode.DateTime:
                                        row[index] = (object)new DateTime(2050, 1, 1);
                                        continue;
                                    case TypeCode.String:
                                        row[index] = (object)"";
                                        continue;
                                    default:
                                        int num = (int)MessageBox.Show("Tipo dato non definito in RecTabRec su " + sTab + "!", "");
                                        continue;
                                }
                            }
                            else
                                row[index] = oleDbDataReader[index];
                        }
                        t.Rows.Add(row);
                        if (bPrimo)
                            break;
                    }
                }
                cm.Dispose();
                oleDbDataReader.Close();
                cn.Dispose();
            }
            catch (Exception ex)
            {
                ErrorLog(ex.Message, sSql);
            }

            return t;
        }

        public string FaiLApice(string strStr)
        {
            string s;
            if (strStr.IndexOfAny(new char[1] { '\'' }) < 0)
            {
                s = strStr;
            }
            else
            {
                string s2 = strStr.Replace("'", "''").Replace("'''", "''");
                Console.WriteLine("aaaaaaaaaa");
                s = s2.Replace("'''", "''").Replace("'''", "''");
                Console.WriteLine("aaaaaaaaaa");
            }

            s = CtrlCrt(s, false);

            return s;
        }

        public string CtrlCrt(string sStr, bool bPun)
        {
            string s = "";
            string sChars = "1234567890qwertyuiopasdfghjklzxcvbnmQWERTYUIOPASDFGHJKLZXCVBNM/'@&.,+-*:%_\\=?!; ";
            if (bPun)
                sChars += ".";

            for (int i = 0; i <= sStr.Length - 1; ++i)
            {
                //string sChar = sStr.Substring(i, 1).ToUpper();
                string sChar = sStr.Substring(i, 1);
                if (sChars.Contains(sChar))
                    s += sChar;
            }
            return s.Trim();
        }

        public string CtrlCrtFile(string sStr)
        {
            string s = "";
            string sChars = "1234567890qwertyuiopasdfghjklzxcvbnmQWERTYUIOPASDFGHJKLZXCVBNM";

            for (int i = 0; i <= sStr.Length - 1; ++i)
            {
                //string sChar = sStr.Substring(i, 1).ToUpper();
                string sChar = sStr.Substring(i, 1);
                if (sChars.Contains(sChar))
                    s += sChar;
            }
            return s.Trim();
        }

        public string CtrlCrt(string sStr, string sCrts)
        {
            string s = "";
            string sChars = sCrts;

            for (int i = 0; i <= sStr.Length - 1; ++i)
            {
                //string sChar = sStr.Substring(i, 1).ToUpper();
                string sChar = sStr.Substring(i, 1);
                if (sChars.Contains(sChar))
                    s += sChar;
            }
            return s.Trim();
        }

        public string DayMdb(DateTime d)
        {
            return "#" + d.ToString("MM-dd-yyyy") + "#";
        }

        public string DaySql(DateTime d)
        {
            return "CONVERT(DATETIME, '" + d.ToString("yyyy-MM-dd") + "', 102)";
        }

        public DateTime Str2Day(string s)
        {
            if (!Numerico(s, "0123456789") || Convert.ToDouble(s) < 010101)
                return this._clsDef.DAYOUT;
            else
            {
                if (s.Length == 6)
                    s = "20" + s;

                return new DateTime((int)Convert.ToInt16(s.Substring(0, 4)), (int)Convert.ToInt16(s.Substring(4, 2)), (int)Convert.ToInt16(s.Substring(6, 2)));
                //return _clsDef.DAYOUT;
            }
        }

        public string StrDayTime(string strOra, string strDay)
        {
            string sOra = "";
            string sDay = "";
            string sRes = "";

            if (strDay.Length >= 10)
            {
                string[] a = strDay.Split('/');

                sDay = a[2] + "-" + a[1] + "-" + a[0];

                if (strDay.Length >= 10)
                    sOra =  "_" + strOra.Replace(":", "-");

                sRes = sDay + sOra;
            }

            return sRes;
        }

        public string DaySel(DateTime d)
        {
            return "#" + d.ToString("yyyy-MM-dd") + "#";
        }

        public int DayDiff(DateTime d1, DateTime d2)
        {
            return Convert.ToInt32((d1 - d2).TotalDays);
        }

        public bool SqlWrite(string sSql, string sDcn)
        {
            bool b = true;
            if (sDcn == "")
                sDcn = ConSql("");
            SqlConnection cn = new SqlConnection(sDcn);
            cn.Open();
            SqlCommand cm = new SqlCommand(sSql, cn);
            try
            {
                cm.ExecuteNonQuery();
            }
            catch (Exception ex)
            {
                ErrorLog(ex.Message, sSql);
                //this._clsMail.InvMail("edp@spacrovigo.org", "EDP", "paolo.secchettin@spacrovigo.org", "", "Errori su SpacServer", "clsGenFun.SqlWrite", "", "");
            }
            cm.Dispose();
            cn.Dispose();
            return b;
        }

        /// <summary>
        /// Esegue una lista di istruzioni SQL su una singola connessione aperta,
        /// riducendo drasticamente i round-trip verso il database rispetto a
        /// chiamate ripetute a SqlWrite().
        /// </summary>
        public bool SqlWriteBatch(List<string> aSql, string sDcn)
        {
            if (aSql == null || aSql.Count == 0) return true;
            bool b = true;
            if (sDcn == "")
                sDcn = ConSql("");
            SqlConnection cn = new SqlConnection(sDcn);
            cn.Open();
            try
            {
                foreach (string sSql in aSql)
                {
                    if (string.IsNullOrEmpty(sSql)) continue;
                    SqlCommand cm = new SqlCommand(sSql, cn);
                    try
                    {
                        cm.ExecuteNonQuery();
                    }
                    catch (Exception ex)
                    {
                        b = false;
                        ErrorLog(ex.Message, sSql);
                    }
                    finally
                    {
                        cm.Dispose();
                    }
                }
            }
            finally
            {
                cn.Dispose();
            }
            return b;
        }


        public bool MdbWrite(string sSql, string sDcn)
        {
            bool b = true;
            if (sDcn == "")
                sDcn = ConMdb("");
            OleDbConnection cn = new OleDbConnection(sDcn);
            cn.Open();
            OleDbCommand cm = new OleDbCommand(sSql, cn);
            try
            {
                cm.ExecuteNonQuery();
            }
            catch (Exception ex)
            {
                ErrorLog(ex.Message, sSql);
                //this._clsMail.InvMail("edp@spacrovigo.org", "EDP", "paolo.secchettin@spacrovigo.org", "", "Errori su SpacServer", "clsGenFun.SqlWrite", "", "");
            }
            cm.Dispose();
            cn.Dispose();
            return b;
        }

        private static readonly object _logLock = new object();

        public static void WriteLogSafe(string filePath, string message)
        {
            if (string.IsNullOrEmpty(filePath)) return;
            System.Threading.ThreadPool.QueueUserWorkItem(_ =>
            {
                try
                {
                    string dir = System.IO.Path.GetDirectoryName(filePath);
                    if (!string.IsNullOrEmpty(dir) && !System.IO.Directory.Exists(dir))
                    {
                        System.IO.Directory.CreateDirectory(dir);
                    }

                    lock (_logLock)
                    {
                        for (int i = 0; i < 3; i++)
                        {
                            try
                            {
                                using (FileStream fs = new FileStream(filePath, FileMode.Append, FileAccess.Write, FileShare.ReadWrite))
                                using (StreamWriter sw = new StreamWriter(fs, System.Text.Encoding.UTF8))
                                {
                                    sw.WriteLine(message);
                                }
                                break;
                            }
                            catch (System.IO.IOException)
                            {
                                if (i == 2) break;
                                System.Threading.Thread.Sleep(30);
                            }
                        }
                    }
                }
                catch { }
            });
        }

        public void ErrorLog(string strSql, string strMsg)
        {
            WriteLogSafe(_clsDef.ERRLOG, DateTime.Now.ToString("dd/MM/yyyy HH.mm") + "; " + strSql + "; " + strMsg + "; ");
        }

        public void FileLog(string strArea, string strCod, string strMsg)
        {
            WriteLogSafe(_clsDef.FILLOG, DateTime.Now.ToString("dd/MM/yyyy HH.mm") + "; " + strArea + "; " + strCod + "; " + strMsg + ";");
        }

        public void FileLogDivFor(string strArea, string strCod, string strMsg)
        {
            WriteLogSafe(_clsDef.FILLOGDIVFOR, DateTime.Now.ToString("dd/MM/yyyy HH.mm") + "; " + strArea + "; " + strCod + "; " + strMsg + ";");
        }

        public void FileFlgDiv(string strArea, string strCod, string strMsg)
        {
            WriteLogSafe(_clsDef.FILLOGDIVFOR, DateTime.Now.ToString("dd/MM/yyyy hh.mm") + "; " + strArea + "; " + strCod + "; " + strMsg + ";");
        }

        public bool Numerico(object o)
        {
            bool flag = true;
            string str1 = "0123456789,.-+";
            if (DBNull.Value.Equals(o))
                flag = false;
            else if (Convert.ToString(o) == "")
            {
                flag = false;
            }
            else
            {
                string str2 = Convert.ToString(o);
                for (int startIndex = 0; startIndex <= str2.Length - 1; ++startIndex)
                {

                    string s = str2.Substring(startIndex, 1);

                    if (str1.IndexOf(s) < 0)
                    {
                        flag = false;
                        break;
                    }
                }
            }
            return flag;
        }

        public string Str2Numerico(string o)
        {
            string res = "";
            string str2 = Convert.ToString(o).Trim();
            string str1 = "0123456789,.-";

            for (int startIndex = 0; startIndex <= str2.Length - 1; ++startIndex)
            {

                string s = str2.Substring(startIndex, 1);

                if (str1.IndexOf(s) > 0)
                {
                    res += s;
                }
            }
           
            return res;
        }

        public bool Numerico(object o, Boolean bolCtrl)
        {
            bool flag = true;
            string str1 = "0123456789";
            if (DBNull.Value.Equals(o))
                flag = false;
            else if (Convert.ToString(o) == "")
            {
                flag = false;
            }
            else
            {
                string str2 = Convert.ToString(o);
                for (int startIndex = 0; startIndex <= str2.Length - 1; ++startIndex)
                {

                    string s = str2.Substring(startIndex, 1);

                    if (str1.IndexOf(s) < 0)
                    {
                        flag = false;
                        break;
                    }
                }
            }
            return flag;
        }

        public bool Numerico(object o, string str1)
        {
            bool flag = true;
            //string str1 = "";
            if (DBNull.Value.Equals(o))
                flag = false;
            else if (Convert.ToString(o) == "")
            {
                flag = false;
            }
            else
            {
                string str2 = Convert.ToString(o);
                for (int startIndex = 0; startIndex <= str2.Length - 1; ++startIndex)
                {

                    string s = str2.Substring(startIndex, 1);

                    if (str1.IndexOf(s) < 0)
                    {
                        flag = false;
                        break;
                    }
                }
            }
            return flag;
        }


        public string Dec2Txt(object decVal, int intDec)
        {
            string s = "0";

            string sDec = "";
            for (int i = 1; i <= intDec; i++)
                sDec += "0";
            if (sDec != "")
                sDec = "#0." + sDec;

                if (!DBNull.Value.Equals(decVal))
                {
                    //s = Convert.ToString(decVal).Replace(",", ".");
                    decimal d = Math.Round(Convert.ToDecimal(decVal), intDec);
                    if(sDec != "")
                        s = d.ToString(sDec);
                    else 
                        s = d.ToString();
                }

            return s;
        }

        public decimal Txt2Dec(string strVal)
        {
            decimal d = 0;
            string s = strVal.Trim();

            if(!DBNull.Value.Equals(strVal) && Numerico(s, "0123456789.,+-"))
            {
                try
                {
                    d = Convert.ToDecimal(Convert.ToString(strVal).Replace(".", ","));
                }
                catch (Exception ex)
                {
                    ErrorLog(ex.Message, strVal);
                    //this._clsMail.InvMail("edp@spacrovigo.org", "EDP", "paolo.secchettin@spacrovigo.org", "", "Errori su SpacServer", "clsGenFun.SqlWrite", "", "");
                }
                //Boolean b = Numerico(strVal);
            }
            return d;
        }

        public decimal Margine(object decVen, object decCes, object decIva, object decSfr, string strTip)
        {
            decimal x = 0;
            decimal d = 0;

            decimal dVen = 0;
            decimal dSfr = 0;
            decimal dCes = 0;
            decimal dIva = 0;

            if (!DBNull.Value.Equals(decVen)) dVen = Convert.ToDecimal(decVen);
            if (!DBNull.Value.Equals(decSfr)) dSfr = Convert.ToDecimal(decSfr);
            if (!DBNull.Value.Equals(decCes)) dCes = Convert.ToDecimal(decCes);
            if (!DBNull.Value.Equals(decIva)) dIva = Convert.ToDecimal(decIva);

            if (dVen > 0 && dCes > 0)
            {
                if (dIva > 0)
                    x = dVen / (1 + (dIva / 100));
                else
                    x = dVen;

                if (dSfr > 0)
                    dCes = dCes * (1 + (dSfr / 100));

                d = x - dCes;

                if (strTip == "P")
                {
                    d = d / x;

                    if (d != 0)
                        d = Math.Round(d * 100, 2);
                }
            }
            else if (dVen > 0 && dCes > 0 && strTip == "P")
            {
                d = (dVen - dCes) / dVen;
                if (d != 0)
                    d = Math.Round(d * 100, 2);

            }

            return d;
        }

        public decimal Marg2Prv(object decMar, object decCes, object decIva, object decSfr)
        {
            decimal d = 0;

            decimal dMar = 0;
            decimal dSfr = 0;
            decimal dCes = 0;
            decimal dIva = 0;

            if (!DBNull.Value.Equals(decMar)) 
                dMar = Convert.ToDecimal(decMar);
            if (!DBNull.Value.Equals(decSfr)) 
                dSfr = Convert.ToDecimal(decSfr);
            if (!DBNull.Value.Equals(decCes)) 
                dCes = Convert.ToDecimal(decCes);
            if (!DBNull.Value.Equals(decIva)) 
                dIva = Convert.ToDecimal(decIva);

            if (dMar > 0 && dCes > 0 && dIva > 0)
            {
                d = (dCes * 100) / (100 - dMar);
                d = PiuIva(d, dIva);
            }

            return d;
        }

        public decimal Ricarico(object decVen, object decCes, object decIva, string strTip)
        {
            decimal x = 0;
            decimal d = 0;

            decimal dVen = 0;
            decimal dCes = 0;
            decimal dIva = 0;

            if (!DBNull.Value.Equals(decVen)) dVen = Convert.ToDecimal(decVen);
            if (!DBNull.Value.Equals(decCes)) dCes = Convert.ToDecimal(decCes);
            if (!DBNull.Value.Equals(decIva)) dIva = Convert.ToDecimal(decIva);

            if (dVen > 0 && dCes > 0 && dIva >= 0)
            {
                x = dIva > 0 ? (dVen / (1 + (dIva / 100))) : dVen;

                d = x - dCes;

                if (strTip == "P")
                    d = x / dCes;

                if (d != 0)
                { 
                    if (strTip == "P")
                        d = Math.Round(d, 2);
                    else
                        d = Math.Round(d, 4);
                }
            }

            return d;
        }

        public decimal MenoIva(decimal decPre, decimal decAli)
        {
            decimal v = decPre;
            if (decAli > 0 && decPre != 0) 
                v = decPre / (1 + (decAli / 100));

            decimal d = decPre - v;

            d = Math.Round(d, 3, MidpointRounding.AwayFromZero);

            v = decPre - d;

            return v;
        }

        public decimal PiuIva(decimal decPre, decimal decAli)
        {
            decimal v = decPre;
            if (decAli > 0 && decPre > 0) v = decPre * (1 + (decAli / 100));
            return v;
        }

        public decimal ValIva(decimal decPre, decimal decAli)
        {
            //decimal v = decPre;
            decimal v = 0;
            if (decAli > 0 && decPre != 0) 
                v = decPre * decAli / 100;
            return v;
        }

        public decimal MenoPer(decimal decVal, decimal decPer)
        {
            Decimal r = decVal;
            Decimal v = decVal;
            Decimal p = decPer;

            if (v > 0 && p > 0)
                r = v - (v * p / 100);

            return Math.Round(r, 2, MidpointRounding.AwayFromZero);
        }

        public string Msg(string strTip, Boolean bolMem, string strMsg)
        {
            string sRes = "";

            frmMsg1 f = new frmMsg1();
            f._strTip = strTip;
            f._strMsg = strMsg;
            f.ShowDialog();
            sRes = f._strRes;

            return sRes;
        }

        public string StrSplit(string str, int len)
        {
            ArrayList aRow = new ArrayList();
            string s = "";

            int iLen = 0;

            string[] a = str.Split(new string[] { "\r\n", "\n" }, StringSplitOptions.None);

            Console.WriteLine("aaaaaaaa");

            iLen = len;

            foreach (string s2 in a)
            {
                string[] aa = s2.Split(' ');

                string sRow = "";

                foreach (string s3 in aa)
                {
                    if (sRow.Length + s3.Length > iLen)
                    {
                        //sRow += new string(' ', iLen);
                        //sRow = sRow.Substring(0, iLen);
                        aRow.Add(sRow);
                        sRow = "";
                    }

                    sRow += s3 + " ";
                }

                if (sRow.Length > 0)
                {
                    aRow.Add(sRow);
                    sRow = "";
                }
            }

            string res = "";

            foreach (string s2 in aRow)
            {

                if (s2.Length <= len)
                    res += s2.PadRight(len, Convert.ToChar(' ')) + "|";
                else if (s2.Length > len)
                    res += s2.Substring(0, len) + "|";
                else
                    res += s2 + "|";
            }

            Console.WriteLine("aaaaaaaaaa");

            return res;
        }

        public string StrSplit2(string str, int len)
        {
            ArrayList aRow = new ArrayList();
            string s = "";

            //int iLen = 0;
            //string[] a = str.Split(new string[] { "\r\n", "\n" }, StringSplitOptions.None);
            //Console.WriteLine("aaaaaaaa");
            //iLen = len;
            //foreach (string s2 in a)
            //{
            //    string[] aa = s2.Split(' ');
            //    string sRow = "";
            //    foreach (string s3 in aa)
            //    {
            //        if (sRow.Length + s3.Length > iLen)
            //        {
            //            //sRow += new string(' ', iLen);
            //            //sRow = sRow.Substring(0, iLen);
            //            aRow.Add(sRow);
            //            sRow = "";
            //        }
            //        sRow += s3 + " ";
            //    }
            //    if (sRow.Length > 0)
            //    {
            //        aRow.Add(sRow);
            //        sRow = "";
            //    }
            //}

            str = str.Replace("\r\n", "");
            str = str.Replace("\n", "");

            string res = "";

            int i = 0;

            len = 50;

            while (true)
            {
                if (str.Length <= i + len)
                {
                    if (str.Length < i + len)
                        res += str.Substring(i) + "|";
                    break;
                }
                else
                {
                    res += str.Substring(i, len) + "|";
                    i += len;
                }

            }

            Console.WriteLine("aaaaaaaaaa");

            return res;
        }

        public void PosRowLast(DataGridView dgv)
        {
            int nRowIndex = dgv.Rows.Count - 1;
            int nColumnIndex = 3;

            dgv.FirstDisplayedScrollingRowIndex = nRowIndex;

            dgv.Rows[nRowIndex].Selected = true;
            dgv.Rows[nRowIndex].Cells[1].Selected = true;
            dgv.CurrentCell = dgv[nColumnIndex, nRowIndex];
        }

        public DateTime FineMese(DateTime dayDay)
        {
            DateTime dDay = dayDay;

            if (dayDay.Month == 12)
                dDay = new DateTime(dayDay.Year + 1, 1, 1);
            else
                dDay = new DateTime(dayDay.Year, dayDay.Month + 1, 1);

            dDay = dDay.AddDays(-1);

            //int iMon = dayDay.Month;

            //while (true)
            //{
            //    dDay = dDay.AddDays(1);
            //    if (dDay.Month > iMon)
            //    {
            //        dDay = dDay.AddDays(-1);
            //        break;
            //    }
            //}

            return dDay;
        }

        //public String DesMese(DateTime dayDay)
        //{
        //    string sRes = String.Format("{0:MMM}", dayDay);

        //    return sRes;
        //}

        public bool DirectoryExists(string directory)
        {
            Func<bool> func = () => Directory.Exists(directory);
            Task<bool> task = new Task<bool>(func);
            task.Start();
            if (task.Wait(500))
            {
                return task.Result;
            }
            else
            {
                // Didn't get an answer back in time be pessimistic and assume it didn't exist
                return false;
            }
        }

        public string PathApOffice(string strPath)
        {
            string sRes = strPath;

            string s = FileIniIni("000");

            if (s != "")
                sRes = sRes.Replace("ApProject", s);

            return sRes;
        }


        //public bool DirectoryExists(string serverName)
        //{
        //    // ***  SET YOUR TIMEOUT HERE  ***     
        //    int timeout = 5;    // 5 seconds 
        //    System.Net.NetworkInformation.Ping pingSender = new System.Net.NetworkInformation.Ping();
        //    System.Net.NetworkInformation.PingOptions options = new System.Net.NetworkInformation.PingOptions();
        //    options.DontFragment = true;
        //    // Enter a valid ip address     
        //    string ipAddressOrHostName = serverName;
        //    string data = "aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa";
        //    byte[] buffer = System.Text.Encoding.ASCII.GetBytes(data);
        //    System.Net.NetworkInformation.PingReply reply = pingSender.Send(ipAddressOrHostName, timeout, buffer, options);
        //    return (reply.Status == System.Net.NetworkInformation.IPStatus.Success);
        //}



    }
}
