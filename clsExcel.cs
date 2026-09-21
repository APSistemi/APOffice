using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Diagnostics;
using System.IO;
using System.Windows.Forms;

namespace APOffice
{
    public class clsExcel
    {
        clsDefine _clsDef = new clsDefine();
        clsFuncs _clsFun = new clsFuncs();

        private string _strPath = "";

        public clsExcel()
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

        public DataTable FillTitStr(string strFld, string strTab)
        {
            DataTable dataTable = new DataTable(strTab);
            if (strTab.Substring(0, 6) == "TabTit")
            {
                dataTable.Columns.Add(new DataColumn()
                {
                    DataType = System.Type.GetType("System.String"),
                    ColumnName = "TitFld",
                    Caption = "Campo",
                    MaxLength = 10,
                    ReadOnly = false
                });
                dataTable.Columns.Add(new DataColumn()
                {
                    DataType = System.Type.GetType("System.String"),
                    ColumnName = "TitDes",
                    Caption = "Descrizione",
                    MaxLength = 30,
                    ReadOnly = false
                });
                dataTable.Columns.Add(new DataColumn()
                {
                    DataType = System.Type.GetType("System.String"),
                    ColumnName = "TitLen",
                    Caption = "Lungo",
                    MaxLength = 3,
                    ReadOnly = false
                });
                dataTable.Columns.Add(new DataColumn()
                {
                    DataType = System.Type.GetType("System.String"),
                    ColumnName = "TitTyp",
                    Caption = "Type",
                    MaxLength = 15,
                    ReadOnly = false
                });
                string str1 = strFld;
                char[] chArray = new char[1]
      {
        ';'
      };
                foreach (string str2 in str1.Split(chArray))
                {
                    if (str2 != "")
                    {
                        string[] strArray = str2.Split(new char[1]
          {
            ','
          });
                        DataRow row = dataTable.NewRow();
                        row["TitFld"] = (object)strArray[0].Trim();
                        row["TitDes"] = (object)strArray[1].Trim();
                        row["TitLen"] = (object)strArray[2].Trim();
                        row["TitTyp"] = (object)strArray[3].Trim();
                        dataTable.Rows.Add(row);
                    }
                }
            }
            else if (strTab == "TabFrmNew")
            {
                dataTable.Columns.Add(new DataColumn()
                {
                    DataType = System.Type.GetType("System.String"),
                    ColumnName = "FrmCod",
                    Caption = "Codice",
                    MaxLength = 5,
                    ReadOnly = false
                });
                dataTable.Columns.Add(new DataColumn()
                {
                    DataType = System.Type.GetType("System.String"),
                    ColumnName = "FrmFrm",
                    Caption = "Formula",
                    MaxLength = 100,
                    ReadOnly = false
                });
                dataTable.Columns.Add(new DataColumn()
                {
                    DataType = System.Type.GetType("System.String"),
                    ColumnName = "FrmCon",
                    Caption = "Campo controllo",
                    MaxLength = 7,
                    ReadOnly = false
                });
                string str1 = strFld;
                char[] chArray = new char[1]
      {
        '?'
      };
                foreach (string str2 in str1.Split(chArray))
                {
                    if (str2 != "")
                    {
                        string[] strArray = str2.Split(new char[1]
          {
            '#'
          });
                        DataRow row = dataTable.NewRow();
                        row["FrmCod"] = (object)strArray[0].Trim();
                        row["FrmFrm"] = (object)strArray[1].Trim();
                        row["FrmCon"] = (object)strArray[2].Trim();
                        dataTable.Rows.Add(row);
                    }
                }
            }
            else if (strTab == "TabFrm")
            {
                dataTable.Columns.Add(new DataColumn()
                {
                    DataType = System.Type.GetType("System.String"),
                    ColumnName = "FrmCod",
                    Caption = "Codice",
                    MaxLength = 5,
                    ReadOnly = false
                });
                dataTable.Columns.Add(new DataColumn()
                {
                    DataType = System.Type.GetType("System.String"),
                    ColumnName = "FrmFrm",
                    Caption = "Formula",
                    MaxLength = 100,
                    ReadOnly = false
                });
                dataTable.Columns.Add(new DataColumn()
                {
                    DataType = System.Type.GetType("System.String"),
                    ColumnName = "FrmCon",
                    Caption = "Campo controllo",
                    MaxLength = 7,
                    ReadOnly = false
                });
                string str1 = strFld;
                char[] chArray = new char[1]
              {
                ';'
              };
                foreach (string str2 in str1.Split(chArray))
                {
                    if (str2 != "")
                    {
                        string[] strArray = str2.Split(new char[1]
                {
                ','
                });
                        DataRow row = dataTable.NewRow();
                        row["FrmCod"] = (object)strArray[0].Trim();
                        row["FrmFrm"] = (object)strArray[1].Trim();
                        row["FrmCon"] = (object)strArray[2].Trim();
                        dataTable.Rows.Add(row);
                    }
                }
            }
            return dataTable;
        }

        public void exportToExcel(DataSet ds, string fileName, string Titolo, string Foot, bool bolVedi)
        {
            DataTable dataTable1 = new DataTable("TabFrmNew");
            foreach (DataTable dataTable2 in (InternalDataCollectionBase)ds.Tables)
            {
                if (dataTable2.TableName == "TabFrmNew")
                {
                    dataTable1 = dataTable2.Copy();
                    ds.Tables.Remove("TabFrmNew");
                    break;
                }
            }
            StreamWriter streamWriter = new StreamWriter(fileName);
            int num1 = 0;
            int num2 = 1;
            streamWriter.Write("<xml version>\r\n<Workbook xmlns=\"urn:schemas-microsoft-com:office:spreadsheet\"\r\n xmlns:o=\"urn:schemas-microsoft-com:office:office\"\r\n xmlns:x=\"urn:schemas-microsoft-com:office:excel\"\r\n xmlns:ss=\"urn:schemas-microsoft-com:office:spreadsheet\">\r\n <Styles>\r\n <Style ss:ID=\"Default\" ss:Name=\"Normal\">\r\n <Alignment ss:Vertical=\"Bottom\"/>\r\n <Borders/>\r\n <Font ss:FontName=\"Arial\" x:Family=\"Swiss\" ss:Size=\"8\"/>\r\n <Interior/>\r\n <NumberFormat/>\r\n <Protection/>\r\n </Style>\r\n <Style ss:ID=\"InCentro\">\r\n <Alignment ss:Horizontal=\"Center\" ss:Vertical=\"Bottom\"/>\r\n </Style>\r\n <Style ss:ID=\"BoldColumn\">\r\n <Font ss:FontName=\"Arial\" x:Family=\"Swiss\" ss:Size=\"8\" ss:Bold=\"1\"/>\r\n </Style>\r\n <Style ss:ID=\"BoldColumn2\">\r\n <Font x:Family=\"Swiss\" ss:Bold=\"1\" ss:Size=\"12\"/>\r\n </Style>\r\n <Style ss:ID=\"StringLiteral\">\r\n <NumberFormat ss:Format=\"@\"/>\r\n </Style>\r\n <Style ss:ID=\"Decimal\">\r\n <NumberFormat ss:Format=\"0.00##\"/>\r\n </Style>\r\n <Style ss:ID=\"Decimal0\">\r\n <NumberFormat ss:Format=\"#,###,###.##\"/>\r\n </Style>\r\n <Style ss:ID=\"Decimal1\">\r\n <NumberFormat ss:Format=\"0.0\"/>\r\n </Style>\r\n <Style ss:ID=\"Decimal2\">\r\n <NumberFormat ss:Format=\"#,###,##0.00\"/>\r\n </Style>\r\n <Style ss:ID=\"Decimal3\">\r\n <NumberFormat ss:Format=\"#,###,##0.000\"/>\r\n </Style>\r\n <Style ss:ID=\"Decimal4\">\r\n <NumberFormat ss:Format=\"0.0000\"/>\r\n </Style>\r\n <Style ss:ID=\"Integer\">\r\n <NumberFormat ss:Format=\"0\"/>\r\n </Style>\r\n <Style ss:ID=\"DateLiteral\">\r\n <NumberFormat ss:Format=\"dd/mm/yyyy;@\"/>\r\n </Style>\r\n <Style ss:ID=\"s67\">\r\n <Alignment ss:Vertical=\"Bottom\" ss:WrapText=\"1\"/>\r\n <Font ss:FontName=\"Arial\" x:Family=\"Swiss\" ss:Size=\"8\" ss:Bold=\"1\"/>\r\n </Style>\r\n </Styles>\r\n ");
            streamWriter.Write("<Worksheet ss:Name=\"Sheet" + (object)num2 + "\">");
            streamWriter.Write("<Table>");
            int num3 = 0;
            foreach (DataRow dataRow in (InternalDataCollectionBase)ds.Tables[0].Rows)
            {
                ++num3;
                streamWriter.Write("<Column ss:Index=\"" + num3.ToString() + "\"  ss:StyleID=\"" + ((string)dataRow["TitTyp"]).Trim() + "\" ss:AutoFitWidth=\"0\" ss:Width=\"" + ((string)dataRow["TitLen"]).Trim() + "\"/>");
            }
            streamWriter.Write("<Row ss:Height=\"15.75\">");
            streamWriter.Write("<Cell ss:StyleID=\"BoldColumn2\"><Data ss:Type=\"String\">" + Titolo + "</Data></Cell>");
            streamWriter.Write("</Row>");
            int index1 = 0;
            while (index1 <= ds.Tables.Count - 1)
            {
                streamWriter.Write("<Row ss:AutoFitHeight=\"0\" ss:Height=\"33.75\" ss:StyleID=\"s67\">");
                foreach (DataRow dataRow in (InternalDataCollectionBase)ds.Tables[index1].Rows)
                {
                    streamWriter.Write("<Cell><Data ss:Type=\"String\">");
                    streamWriter.Write(dataRow["TitDes"]);
                    streamWriter.Write("</Data></Cell>");
                }
                streamWriter.Write("</Row>");
                foreach (DataRow dataRow1 in (InternalDataCollectionBase)ds.Tables[index1 + 1].Rows)
                {
                    int num4 = 0;
                    ++num1;
                    if (num1 == 64000)
                    {
                        num1 = 0;
                        ++num2;
                        streamWriter.Write("</Table>");
                        streamWriter.Write(" </Worksheet>");
                        streamWriter.Write("<Worksheet ss:Name=\"Sheet" + (object)num2 + "\">");
                        streamWriter.Write("<Table>");
                    }
                    streamWriter.Write("<Row>");
                    foreach (DataRow dataRow2 in (InternalDataCollectionBase)ds.Tables[index1].Rows)
                    {
                        string index2 = (string)dataRow2["TitFld"];
                        ++num4;
                        if (index2.Length >= 3 && index2.Substring(0, 3).ToUpper() == "FRM")
                        {
                            DataRow[] dataRowArray = dataTable1.Select("FrmCod='" + index2 + "'");
                            if (dataRowArray.Length > 0)
                            {
                                bool flag = true;
                                if ((string)dataRowArray[0]["FrmCon"] == "NO")
                                    flag = false;
                                else if (DBNull.Value.Equals(dataRow1[(string)dataRowArray[0]["FrmCon"]]) || (Decimal)dataRow1[(string)dataRowArray[0]["FrmCon"]] == new Decimal(0))
                                    flag = false;
                                if (!flag)
                                {
                                    streamWriter.Write("<Cell ss:StyleID=\"StringLiteral\"><Data ss:Type=\"String\"></Data></Cell>");
                                }
                                else
                                {
                                    string str = (string)dataRowArray[0]["FrmFrm"];
                                    streamWriter.Write("<Cell ss:StyleID=\"" + ((string)dataRow2["TitTyp"]).Trim() + "\" ss:Formula=\"=" + str + "\"><Data ss:Type=\"Number\">0</Data></Cell>");
                                }
                            }
                        }
                        else
                        {
                            Type type = dataRow1[index2].GetType();
                            switch (type.ToString())
                            {
                                case "System.String":
                                    string str1 = dataRow1[index2].ToString();
                                    if (str1 == "")
                                    {
                                        streamWriter.Write("<Cell ss:StyleID=\"Decimal0\"></Cell>");
                                        break;
                                    }
                                    else
                                    {
                                        string str2 = str1.Trim().Replace("&", "&").Replace(">", "").Replace("<", "");
                                        streamWriter.Write("<Cell ss:StyleID=\"StringLiteral\"><Data ss:Type=\"String\">");
                                        streamWriter.Write(str2);
                                        streamWriter.Write("</Data></Cell>");
                                        break;
                                    }
                                case "System.DateTime":
                                    DateTime dateTime = (DateTime)dataRow1[index2];
                                    int num5 = dateTime.Year;
                                    if (num5.ToString() == "2050")
                                    {
                                        streamWriter.Write("<Cell ss:StyleID=\"StringLiteral\"><Data ss:Type=\"String\"></Data></Cell>");
                                        break;
                                    }
                                    else
                                    {
                                        string[] strArray1 = new string[12];
                                        string[] strArray2 = strArray1;
                                        int index3 = 0;
                                        num5 = dateTime.Year;
                                        string str2 = num5.ToString();
                                        strArray2[index3] = str2;
                                        strArray1[1] = "-";
                                        string[] strArray3 = strArray1;
                                        int index4 = 2;
                                        string str3;
                                        if (dateTime.Month >= 10)
                                        {
                                            num5 = dateTime.Month;
                                            str3 = num5.ToString();
                                        }
                                        else
                                        {
                                            string str4 = "0";
                                            num5 = dateTime.Month;
                                            string str5 = num5.ToString();
                                            str3 = str4 + str5;
                                        }
                                        strArray3[index4] = str3;
                                        strArray1[3] = "-";
                                        string[] strArray4 = strArray1;
                                        int index5 = 4;
                                        string str6;
                                        if (dateTime.Day >= 10)
                                        {
                                            num5 = dateTime.Day;
                                            str6 = num5.ToString();
                                        }
                                        else
                                        {
                                            string str4 = "0";
                                            num5 = dateTime.Day;
                                            string str5 = num5.ToString();
                                            str6 = str4 + str5;
                                        }
                                        strArray4[index5] = str6;
                                        strArray1[5] = "T";
                                        string[] strArray5 = strArray1;
                                        int index6 = 6;
                                        string str7;
                                        if (dateTime.Hour >= 10)
                                        {
                                            num5 = dateTime.Hour;
                                            str7 = num5.ToString();
                                        }
                                        else
                                        {
                                            string str4 = "0";
                                            num5 = dateTime.Hour;
                                            string str5 = num5.ToString();
                                            str7 = str4 + str5;
                                        }
                                        strArray5[index6] = str7;
                                        strArray1[7] = ":";
                                        string[] strArray6 = strArray1;
                                        int index7 = 8;
                                        string str8;
                                        if (dateTime.Minute >= 10)
                                        {
                                            num5 = dateTime.Minute;
                                            str8 = num5.ToString();
                                        }
                                        else
                                        {
                                            string str4 = "0";
                                            num5 = dateTime.Minute;
                                            string str5 = num5.ToString();
                                            str8 = str4 + str5;
                                        }
                                        strArray6[index7] = str8;
                                        strArray1[9] = ":";
                                        string[] strArray7 = strArray1;
                                        int index8 = 10;
                                        string str9;
                                        if (dateTime.Second >= 10)
                                        {
                                            num5 = dateTime.Second;
                                            str9 = num5.ToString();
                                        }
                                        else
                                        {
                                            string str4 = "0";
                                            num5 = dateTime.Second;
                                            string str5 = num5.ToString();
                                            str9 = str4 + str5;
                                        }
                                        strArray7[index8] = str9;
                                        strArray1[11] = ".000";
                                        string str10 = string.Concat(strArray1);
                                        streamWriter.Write("<Cell ss:StyleID=\"DateLiteral\"><Data ss:Type=\"DateTime\">");
                                        streamWriter.Write(str10);
                                        streamWriter.Write("</Data></Cell>");
                                        break;
                                    }
                                case "System.Boolean":
                                    streamWriter.Write("<Cell ss:StyleID=\"StringLiteral\"><Data ss:Type=\"String\">");
                                    streamWriter.Write(dataRow1[index2].ToString());
                                    streamWriter.Write("</Data></Cell>");
                                    break;
                                case "System.Int16":
                                case "System.Int32":
                                case "System.Int64":
                                case "System.Byte":
                                    streamWriter.Write("<Cell ss:StyleID=\"Integer\"><Data ss:Type=\"Number\">");
                                    streamWriter.Write(dataRow1[index2].ToString());
                                    streamWriter.Write("</Data></Cell>");
                                    break;
                                case "System.Decimal":
                                case "System.Double":
                                    string str11 = dataRow1[index2].ToString();
                                    if (Convert.ToDecimal(str11) == new Decimal(0))
                                    {
                                        streamWriter.Write("<Cell ss:StyleID=\"Decimal0\"></Cell>");
                                        break;
                                    }
                                    else
                                    {
                                        streamWriter.Write("<Cell><Data ss:Type=\"Number\">");
                                        string str2 = str11.Replace(",", ".");
                                        streamWriter.Write(str2);
                                        streamWriter.Write("</Data></Cell>");
                                        break;
                                    }
                                case "System.DBNull":
                                    streamWriter.Write("<Cell ss:StyleID=\"Decimal0\"></Cell>");
                                    break;
                                default:
                                    throw new Exception(type.ToString() + " not handled.");
                            }
                        }
                    }
                    streamWriter.Write("</Row>");
                }
                index1 += 2;
            }
            if (Foot != "")
            {

                string s = "";
                int iCol = 0;
                string[] a = Foot.Split('|');

                if (a.Length > 1)
                {
                    //if (_clsFun.Numerico(a[0]))

                    if (int.TryParse(a[0], out iCol))
                    {
                        iCol = Convert.ToInt16(a[0]);
                        s = Convert.ToString(a[1]);
                    }
                    else
                        s = Foot;
                }
                else
                    s = Foot;


                streamWriter.Write("<Row ss:Height=\"15.75\">");
                for (int i = 0; i < iCol; i++)
                    streamWriter.Write("<Cell></Cell>");
                streamWriter.Write("<Cell ss:StyleID=\"BoldColumn2\"><Data ss:Type=\"String\">" + s + "</Data></Cell>");
                streamWriter.Write("</Row>");
                //}

                //streamWriter.Write("<Row ss:Height=\"15.75\">");
                //streamWriter.Write("<Cell ss:StyleID=\"BoldColumn2\"><Data ss:Type=\"String\">" + Foot + "</Data></Cell>");
                //streamWriter.Write("</Row>");
            }
            streamWriter.Write("</Table>");
            streamWriter.Write(" </Worksheet>");
            streamWriter.Write("</Workbook>");
            streamWriter.Close();
            if (!bolVedi)
                return;
            this.Processo(fileName);
        }

        private void exportToExcel(DataSet source, string fileName)
        {

            System.IO.StreamWriter excelDoc;

            excelDoc = new System.IO.StreamWriter(fileName);
            const string startExcelXML = "<xml version>\r\n<Workbook " +
                    "xmlns=\"urn:schemas-microsoft-com:office:spreadsheet\"\r\n" +
                    " xmlns:o=\"urn:schemas-microsoft-com:office:office\"\r\n " +
                    "xmlns:x=\"urn:schemas-    microsoft-com:office:" +
                    "excel\"\r\n xmlns:ss=\"urn:schemas-microsoft-com:" +
                    "office:spreadsheet\">\r\n <Styles>\r\n " +
                    "<Style ss:ID=\"Default\" ss:Name=\"Normal\">\r\n " +
                    "<Alignment ss:Vertical=\"Bottom\"/>\r\n <Borders/>" +
                    "\r\n <Font/>\r\n <Interior/>\r\n <NumberFormat/>" +
                    "\r\n <Protection/>\r\n </Style>\r\n " +
                    "<Style ss:ID=\"BoldColumn\">\r\n <Font " +
                    "x:Family=\"Swiss\" ss:Bold=\"1\"/>\r\n </Style>\r\n " +
                    "<Style     ss:ID=\"StringLiteral\">\r\n <NumberFormat" +
                    " ss:Format=\"@\"/>\r\n </Style>\r\n <Style " +
                    "ss:ID=\"Decimal\">\r\n <NumberFormat " +
                    "ss:Format=\"0.00##\"/>\r\n </Style>\r\n " +
                    "<Style ss:ID=\"Integer\">\r\n <NumberFormat " +
                    "ss:Format=\"0\"/>\r\n </Style>\r\n <Style " +
                    "ss:ID=\"DateLiteral\">\r\n <NumberFormat " +
                    "ss:Format=\"mm/dd/yyyy;@\"/>\r\n </Style>\r\n " +
                    "</Styles>\r\n ";
            const string endExcelXML = "</Workbook>";

            int rowCount = 0;
            int sheetCount = 1;
            /*
            <xml version>
            <Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet"
            xmlns:o="urn:schemas-microsoft-com:office:office"
            xmlns:x="urn:schemas-microsoft-com:office:excel"
            xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet">
            <Styles>
            <Style ss:ID="Default" ss:Name="Normal">
                <Alignment ss:Vertical="Bottom"/>
                <Borders/>
                <Font/>
                <Interior/>
                <NumberFormat/>
                <Protection/>
            </Style>
            <Style ss:ID="BoldColumn">
                <Font x:Family="Swiss" ss:Bold="1"/>
            </Style>
            <Style ss:ID="StringLiteral">
                <NumberFormat ss:Format="@"/>
            </Style>
            <Style ss:ID="Decimal">
                <NumberFormat ss:Format="0.0000"/>
            </Style>
            <Style ss:ID="Integer">
                <NumberFormat ss:Format="0"/>
            </Style>
            <Style ss:ID="DateLiteral">
                <NumberFormat ss:Format="mm/dd/yyyy;@"/>
            </Style>
            </Styles>
            <Worksheet ss:Name="Sheet1">
            </Worksheet>
            </Workbook>
            */
            excelDoc.Write(startExcelXML);
            excelDoc.Write("<Worksheet ss:Name=\"Sheet" + sheetCount + "\">");
            excelDoc.Write("<Table>");
            excelDoc.Write("<Row>");
            //for (int x = 0; x < source.Tables[0].Rows.Count; x++)
            //{
            //    excelDoc.Write("<Cell ss:StyleID=\"BoldColumn\"><Data ss:Type=\"String\">");
            //    //excelDoc.Write(source.Tables[0].Columns[x].ColumnName);
            //    excelDoc.Write(source.Tables[0].Columns[x].Caption);
            //    excelDoc.Write("</Data></Cell>");
            //}

            foreach (DataRow x in source.Tables[0].Rows)
            {
                excelDoc.Write("<Cell ss:StyleID=\"BoldColumn\"><Data ss:Type=\"String\">");
                //excelDoc.Write(source.Tables[0].Columns[x].ColumnName);
                excelDoc.Write(x["TitDes"]);
                excelDoc.Write("</Data></Cell>");
            }

            excelDoc.Write("</Row>");
            foreach (DataRow x in source.Tables[1].Rows)
            {
                rowCount++;
                //if the number of rows is > 64000 create a new page to continue output
                if (rowCount == 64000)
                {
                    rowCount = 0;
                    sheetCount++;
                    excelDoc.Write("</Table>");
                    excelDoc.Write(" </Worksheet>");
                    excelDoc.Write("<Worksheet ss:Name=\"Sheet" + sheetCount + "\">");
                    excelDoc.Write("<Table>");
                }
                excelDoc.Write("<Row>"); //ID=" + rowCount + "
                for (int y = 0; y < source.Tables[1].Columns.Count; y++)
                {
                    System.Type rowType;
                    rowType = x[y].GetType();
                    switch (rowType.ToString())
                    {
                        case "System.String":
                            string XMLstring = x[y].ToString();
                            XMLstring = XMLstring.Trim();
                            XMLstring = XMLstring.Replace("&", "&");
                            XMLstring = XMLstring.Replace(">", ">");
                            XMLstring = XMLstring.Replace("<", "<");
                            excelDoc.Write("<Cell ss:StyleID=\"StringLiteral\">" +
                                            "<Data ss:Type=\"String\">");
                            excelDoc.Write(XMLstring);
                            excelDoc.Write("</Data></Cell>");
                            break;
                        case "System.DateTime":
                            //Excel has a specific Date Format of YYYY-MM-DD followed by  
                            //the letter 'T' then hh:mm:sss.lll Example 2005-01-31T24:01:21.000
                            //The Following Code puts the date stored in XMLDate 
                            //to the format above
                            DateTime XMLDate = (DateTime)x[y];
                            string XMLDatetoString = ""; //Excel Converted Date
                            XMLDatetoString = XMLDate.Year.ToString() +
                                    "-" +
                                    (XMLDate.Month < 10 ? "0" +
                                    XMLDate.Month.ToString() : XMLDate.Month.ToString()) +
                                    "-" +
                                    (XMLDate.Day < 10 ? "0" +
                                    XMLDate.Day.ToString() : XMLDate.Day.ToString()) +
                                    "T" +
                                    (XMLDate.Hour < 10 ? "0" +
                                    XMLDate.Hour.ToString() : XMLDate.Hour.ToString()) +
                                    ":" +
                                    (XMLDate.Minute < 10 ? "0" +
                                    XMLDate.Minute.ToString() : XMLDate.Minute.ToString()) +
                                    ":" +
                                    (XMLDate.Second < 10 ? "0" +
                                    XMLDate.Second.ToString() : XMLDate.Second.ToString()) +
                                    ".000";
                            excelDoc.Write("<Cell ss:StyleID=\"DateLiteral\">" +
                                            "<Data ss:Type=\"DateTime\">");
                            excelDoc.Write(XMLDatetoString);
                            excelDoc.Write("</Data></Cell>");
                            break;
                        case "System.Boolean":
                            excelDoc.Write("<Cell ss:StyleID=\"StringLiteral\">" +
                                        "<Data ss:Type=\"String\">");
                            excelDoc.Write(x[y].ToString());
                            excelDoc.Write("</Data></Cell>");
                            break;
                        case "System.Int16":
                        case "System.Int32":
                        case "System.Int64":
                        case "System.Byte":
                            excelDoc.Write("<Cell ss:StyleID=\"Integer\">" +
                                    "<Data ss:Type=\"Number\">");
                            excelDoc.Write(x[y].ToString());
                            excelDoc.Write("</Data></Cell>");
                            break;
                        case "System.Decimal":
                        case "System.Double":
                            excelDoc.Write("<Cell ss:StyleID=\"Decimal\">" +
                                    "<Data ss:Type=\"Number\">");
                            string s = x[y].ToString();
                            s = s.Replace(",", ".");
                            excelDoc.Write(s);
                            excelDoc.Write("</Data></Cell>");
                            break;
                        case "System.DBNull":
                            excelDoc.Write("<Cell ss:StyleID=\"StringLiteral\">" +
                                    "<Data ss:Type=\"String\">");
                            excelDoc.Write("");
                            excelDoc.Write("</Data></Cell>");
                            break;
                        default:
                            throw (new Exception(rowType.ToString() + " not handled."));
                    }
                }
                excelDoc.Write("</Row>");
            }
            excelDoc.Write("</Table>");
            excelDoc.Write(" </Worksheet>");
            excelDoc.Write(endExcelXML);
            excelDoc.Close();
        }

        public void exportToExcel1(DataSet ds, string fileName, string Titolo, string Foot)
        {

            //DataSet ds = new DataSet();
            DataTable tFrm = new DataTable("TabFrm");

            foreach (DataTable t in ds.Tables)
            {
                if (t.TableName == "TabFrm")
                {
                    tFrm = t.Copy();
                    ds.Tables.Remove("TabFrm");
                    break;
                }
                //else
                //    ds.Tables.Add(t);
            }

            System.IO.StreamWriter excelDoc;
            string s = "";

            //<Style ss:ID="s67">
            // <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>
            //</Style>

            excelDoc = new System.IO.StreamWriter(fileName);
            const string startExcelXML =
                    "<xml version>\r\n" +
                    "<Workbook xmlns=\"urn:schemas-microsoft-com:office:spreadsheet\"\r\n " +
                    "xmlns:o=\"urn:schemas-microsoft-com:office:office\"\r\n " +
                    "xmlns:x=\"urn:schemas-microsoft-com:office:excel\"\r\n " +
                    "xmlns:ss=\"urn:schemas-microsoft-com:office:spreadsheet\">\r\n " +
                    "<Styles>\r\n " +
                    "<Style ss:ID=\"Default\" ss:Name=\"Normal\">\r\n " +
                        "<Alignment ss:Vertical=\"Bottom\"/>\r\n <Borders/>\r\n " +
                        "<Font ss:FontName=\"Arial\" x:Family=\"Swiss\" ss:Size=\"8\"/>\r\n " +
                        "<Interior/>\r\n " +
                        "<NumberFormat/>\r\n " +
                        "<Protection/>\r\n " +
                    "</Style>\r\n " +
                    "<Style ss:ID=\"InCentro\">\r\n " +
                        "<Alignment ss:Horizontal=\"Center\" ss:Vertical=\"Bottom\"/>\r\n " +
                    "</Style>\r\n " +
                    "<Style ss:ID=\"BoldColumn\">\r\n " +
                        "<Font ss:FontName=\"Arial\" x:Family=\"Swiss\" ss:Size=\"8\" ss:Bold=\"1\"/>\r\n " +
                    "</Style>\r\n " +
                    "<Style ss:ID=\"BoldColumn2\">\r\n " +
                        "<Font x:Family=\"Swiss\" ss:Bold=\"1\" ss:Size=\"12\"/>\r\n " +
                    "</Style>\r\n " +
                    "<Style ss:ID=\"StringLiteral\">\r\n " +
                        "<NumberFormat ss:Format=\"@\"/>\r\n " +
                    "</Style>\r\n " +
                    "<Style ss:ID=\"Decimal\">\r\n " +
                        "<NumberFormat ss:Format=\"0.00##\"/>\r\n " +
                    "</Style>\r\n " +
                    "<Style ss:ID=\"Decimal0\">\r\n " +
                        "<NumberFormat ss:Format=\"#,###,###.##\"/>\r\n " +
                    "</Style>\r\n " +
                    "<Style ss:ID=\"Decimal1\">\r\n " +
                        "<NumberFormat ss:Format=\"0.0\"/>\r\n " +
                    "</Style>\r\n " +
                    "<Style ss:ID=\"Decimal2\">\r\n " +
                        "<NumberFormat ss:Format=\"0.00\"/>\r\n " +
                    "</Style>\r\n " +
                    "<Style ss:ID=\"Decimal3\">\r\n " +
                        "<NumberFormat ss:Format=\"0.000\"/>\r\n " +
                    "</Style>\r\n " +
                    "<Style ss:ID=\"Decimal4\">\r\n " +
                        "<NumberFormat ss:Format=\"0.0000\"/>\r\n " +
                    "</Style>\r\n " +
                    "<Style ss:ID=\"Integer\">\r\n " +
                        "<NumberFormat ss:Format=\"0\"/>\r\n " +
                    "</Style>\r\n " +
                    "<Style ss:ID=\"DateLiteral\">\r\n " +
                        "<NumberFormat ss:Format=\"dd/mm/yyyy;@\"/>\r\n " +
                    "</Style>\r\n " +
                    "<Style ss:ID=\"s67\">\r\n " +
                        "<Alignment ss:Vertical=\"Bottom\" ss:WrapText=\"1\"/>\r\n " +
                        "<Font ss:FontName=\"Arial\" x:Family=\"Swiss\" ss:Size=\"8\" ss:Bold=\"1\"/>\r\n " +
                        "</Style>\r\n " +
                    "</Styles>\r\n ";
            const string endExcelXML = "</Workbook>";

            int colCount = 0;
            int rowCount = 0;
            int sheetCount = 1;
            //int i = 1;
            DataRow[] j;
            excelDoc.Write(startExcelXML);
            excelDoc.Write("<Worksheet ss:Name=\"Sheet" + sheetCount + "\">");
            excelDoc.Write("<Table>");

            int iL = 0;
            foreach (DataRow x in ds.Tables[0].Rows)
            {
                iL += 1;
                excelDoc.Write("<Column ss:Index=\"" + iL.ToString() + "\"  ss:StyleID=\"" + ((string)x["TitTyp"]).Trim() + "\" ss:AutoFitWidth=\"0\" ss:Width=\"" + ((string)x["TitLen"]).Trim() + "\"/>");
            }

            excelDoc.Write("<Row ss:Height=\"15.75\">");
            excelDoc.Write("<Cell ss:StyleID=\"BoldColumn2\"><Data ss:Type=\"String\">" + Titolo + "</Data></Cell>");
            excelDoc.Write("</Row>");

            /**/

            for (int i = 0; i <= ds.Tables.Count - 1; i += 2)
            {

                excelDoc.Write("<Row ss:AutoFitHeight=\"0\" ss:Height=\"33.75\" ss:StyleID=\"s67\">");
                foreach (DataRow x in ds.Tables[i].Rows)
                {
                    excelDoc.Write("<Cell><Data ss:Type=\"String\">");
                    excelDoc.Write(x["TitDes"]);
                    excelDoc.Write("</Data></Cell>");
                }
                excelDoc.Write("</Row>");

                foreach (DataRow x in ds.Tables[i + 1].Rows)
                {

                    colCount = 0;
                    rowCount++;
                    //string s2 = "";
                    //if the number of rows is > 64000 create a new page to continue output
                    if (rowCount == 64000)
                    {
                        rowCount = 0;
                        sheetCount++;
                        excelDoc.Write("</Table>");
                        excelDoc.Write(" </Worksheet>");
                        excelDoc.Write("<Worksheet ss:Name=\"Sheet" + sheetCount + "\">");
                        excelDoc.Write("<Table>");
                    }
                    excelDoc.Write("<Row>");
                    //for (int y = 0; y < source.Tables[1].Columns.Count; y++)
                    foreach (DataRow y in ds.Tables[i].Rows)
                    {

                        string sFld = (string)y["TitFld"];

                        //s2 += source.Tables[1].Columns[y].ColumnName + " | ";
                        //j = source.Tables[0].Select("TitFld='" + source.Tables[1].Columns[y].ColumnName + "'");
                        //if (j.Length > 0)
                        //{
                        colCount++;

                        if (sFld.Substring(0, 3).ToUpper() == "FRM")
                        {
                            j = tFrm.Select("FrmCod='" + sFld + "'");
                            if (j.Length > 0)
                            {
                                if ((decimal)x[(string)j[0]["FrmCon"]] == 0)
                                    excelDoc.Write("<Cell ss:StyleID=\"StringLiteral\"><Data ss:Type=\"String\"></Data></Cell>");
                                else
                                    excelDoc.Write("<Cell ss:StyleID=\"" + ((string)y["TitTyp"]).Trim() + "\" ss:Formula=\"=" + (string)j[0]["FrmFrm"] + "\"><Data ss:Type=\"Number\">0</Data></Cell>");
                                //excelDoc.Write("<Cell ss:StyleID=\"Decimal1\" ss:Formula=\"=" + (string)j[0]["FrmFrm"] + "\"><Data ss:Type=\"Number\">0</Data></Cell>");

                            }

                            //if ((string)x["pam_art"] == "")
                            //    excelDoc.Write("<Cell ss:StyleID=\"StringLiteral\"><Data ss:Type=\"String\"></Data></Cell>");
                            //else
                            //{
                            //    if (sFld == "Form1")
                            //        s = "((RC[-1]-RC[-2])/RC[-1])*100";
                            //    if (sFld == "Form2")
                            //        s = "(RC[-2]-(RC[-5]+(RC[-5]*RC[-15])/100))/RC[-2]*100";
                            //    if (sFld == "Form3")
                            //        s = "(RC[-4]-(RC[-8]+(RC[-8]*RC[-17])/100))/RC[-4]*100";
                            //    if (sFld == "Form4")
                            //        s = "RC[-10]+(RC[-1]*RC[-10]/100)";
                            //    if (sFld == "Form5")
                            //        s = "(RC[+1]-RC[-1])/RC[+1]*100";
                            //    if (sFld == "Form6")
                            //        s = "RC[-15]-(RC[-15]*RC[+1]/100)";
                            //    if (sFld == "Form7")
                            //        s = "RC[-22]/(100 + RC[-22])*100";
                            //    //s = "RC[-21]/(100+RC[-21])*100)";
                            //    excelDoc.Write("<Cell ss:StyleID=\"Decimal\" ss:Formula=\"=" + s + "\"><Data ss:Type=\"Number\">0</Data></Cell>");
                            //}

                        }

                        else
                        {

                            System.Type rowType;
                            rowType = x[sFld].GetType();
                            switch (rowType.ToString())
                            {
                                case "System.String":
                                    s = x[sFld].ToString();
                                    if (s == "")
                                    {
                                        excelDoc.Write("<Cell ss:StyleID=\"Decimal0\"></Cell>");
                                    }
                                    else
                                    {
                                        s = s.Trim();
                                        s = s.Replace("&", "&");
                                        s = s.Replace(">", "");
                                        s = s.Replace("<", "");
                                        //if (ds.Tables[1].Columns[sFld].ColumnName == "fat_cli")
                                        //    excelDoc.Write("<Cell ss:StyleID=\"InCentro\"><Data ss:Type=\"String\">");
                                        //else
                                        excelDoc.Write("<Cell ss:StyleID=\"StringLiteral\"><Data ss:Type=\"String\">");
                                        excelDoc.Write(s);
                                        excelDoc.Write("</Data></Cell>");
                                    }
                                    break;
                                case "System.DateTime":
                                    //Excel has a specific Date Format of YYYY-MM-DD followed by  
                                    //the letter 'T' then hh:mm:sss.lll Example 2005-01-31T24:01:21.000
                                    //The Following Code puts the date stored in XMLDate 
                                    //to the format above
                                    DateTime XMLDate = (DateTime)x[sFld];
                                    string XMLDatetoString = ""; //Excel Converted Date
                                    if (XMLDate.Year.ToString() == "2050")
                                        excelDoc.Write("<Cell ss:StyleID=\"StringLiteral\"><Data ss:Type=\"String\"></Data></Cell>");
                                    else
                                    {
                                        XMLDatetoString = XMLDate.Year.ToString() +
                                                "-" +
                                                (XMLDate.Month < 10 ? "0" +
                                                XMLDate.Month.ToString() : XMLDate.Month.ToString()) +
                                                "-" +
                                                (XMLDate.Day < 10 ? "0" +
                                                XMLDate.Day.ToString() : XMLDate.Day.ToString()) +
                                                "T" +
                                                (XMLDate.Hour < 10 ? "0" +
                                                XMLDate.Hour.ToString() : XMLDate.Hour.ToString()) +
                                                ":" +
                                                (XMLDate.Minute < 10 ? "0" +
                                                XMLDate.Minute.ToString() : XMLDate.Minute.ToString()) +
                                                ":" +
                                                (XMLDate.Second < 10 ? "0" +
                                                XMLDate.Second.ToString() : XMLDate.Second.ToString()) +
                                                ".000";
                                        excelDoc.Write("<Cell ss:StyleID=\"DateLiteral\"><Data ss:Type=\"DateTime\">");
                                        excelDoc.Write(XMLDatetoString);
                                        excelDoc.Write("</Data></Cell>");
                                    }
                                    break;
                                case "System.Boolean":
                                    excelDoc.Write("<Cell ss:StyleID=\"StringLiteral\"><Data ss:Type=\"String\">");
                                    excelDoc.Write(x[sFld].ToString());
                                    excelDoc.Write("</Data></Cell>");
                                    break;
                                case "System.Int16":
                                case "System.Int32":
                                case "System.Int64":
                                case "System.Byte":
                                    excelDoc.Write("<Cell ss:StyleID=\"Integer\">" +
                                            "<Data ss:Type=\"Number\">");
                                    excelDoc.Write(x[sFld].ToString());
                                    excelDoc.Write("</Data></Cell>");
                                    break;
                                case "System.Decimal":
                                case "System.Double":
                                    s = x[sFld].ToString();
                                    if (Convert.ToDecimal(s) == 0)
                                        excelDoc.Write("<Cell ss:StyleID=\"Decimal0\"></Cell>");
                                    else
                                    {
                                        excelDoc.Write("<Cell><Data ss:Type=\"Number\">");
                                        s = s.Replace(",", ".");
                                        excelDoc.Write(s);
                                        excelDoc.Write("</Data></Cell>");
                                    }
                                    break;
                                case "System.DBNull":
                                    excelDoc.Write("<Cell ss:StyleID=\"StringLiteral\"><Data ss:Type=\"String\">");
                                    excelDoc.Write("");
                                    excelDoc.Write("</Data></Cell>");
                                    break;
                                default:
                                    throw (new Exception(rowType.ToString() + " not handled."));
                            }

                        }
                    }

                    excelDoc.Write("</Row>");

                }

            }

            if (Foot != "")
            {
                excelDoc.Write("<Row ss:Height=\"15.75\">");
                excelDoc.Write("<Cell ss:StyleID=\"BoldColumn2\"><Data ss:Type=\"String\">" + Foot + "</Data></Cell>");
                excelDoc.Write("</Row>");
            }

            excelDoc.Write("</Table>");
            excelDoc.Write(" </Worksheet>");
            excelDoc.Write(endExcelXML);
            excelDoc.Close();

            Processo(fileName);

        }

        public void exportToExcel1Old(DataSet ds, string fileName, string Titolo, string Foot)
        {

            System.IO.StreamWriter excelDoc;
            string s = "";

            excelDoc = new System.IO.StreamWriter(fileName);
            const string startExcelXML =
                    "<xml version>\r\n" +
                    "<Workbook xmlns=\"urn:schemas-microsoft-com:office:spreadsheet\"\r\n " +
                    "xmlns:o=\"urn:schemas-microsoft-com:office:office\"\r\n " +
                    "xmlns:x=\"urn:schemas-microsoft-com:office:excel\"\r\n " +
                    "xmlns:ss=\"urn:schemas-microsoft-com:office:spreadsheet\">\r\n " +
                    "<Styles>\r\n " +
                    "<Style ss:ID=\"Default\" ss:Name=\"Normal\">\r\n " +
                        "<Alignment ss:Vertical=\"Bottom\"/>\r\n <Borders/>\r\n " +
                        "<Font ss:FontName=\"Arial\" x:Family=\"Swiss\" ss:Size=\"8\"/>\r\n " +
                        "<Interior/>\r\n " +
                        "<NumberFormat/>\r\n " +
                        "<Protection/>\r\n " +
                    "</Style>\r\n " +
                    "<Style ss:ID=\"InCentro\">\r\n " +
                        "<Alignment ss:Horizontal=\"Center\" ss:Vertical=\"Bottom\"/>\r\n " +
                    "</Style>\r\n " +
                    "<Style ss:ID=\"BoldColumn\">\r\n " +
                        "<Font ss:FontName=\"Arial\" x:Family=\"Swiss\" ss:Size=\"8\" ss:Bold=\"1\"/>\r\n " +
                    "</Style>\r\n " +
                    "<Style ss:ID=\"BoldColumn2\">\r\n " +
                        "<Font x:Family=\"Swiss\" ss:Bold=\"1\" ss:Size=\"12\"/>\r\n " +
                    "</Style>\r\n " +
                    "<Style ss:ID=\"StringLiteral\">\r\n " +
                        "<NumberFormat ss:Format=\"@\"/>\r\n " +
                    "</Style>\r\n " +
                    "<Style ss:ID=\"Decimal\">\r\n " +
                        "<NumberFormat ss:Format=\"0.00##\"/>\r\n " +
                    "</Style>\r\n " +
                    "<Style ss:ID=\"Decimal2\">\r\n " +
                        "<NumberFormat ss:Format=\"0.00\"/>\r\n " +
                    "</Style>\r\n " +
                    "<Style ss:ID=\"Integer\">\r\n " +
                        "<NumberFormat ss:Format=\"0\"/>\r\n " +
                    "</Style>\r\n " +
                    "<Style ss:ID=\"DateLiteral\">\r\n " +
                        "<NumberFormat ss:Format=\"dd/mm/yyyy;@\"/>\r\n " +
                    "</Style>\r\n " +
                    "<Style ss:ID=\"s67\">\r\n " +
                        "<Alignment ss:Vertical=\"Bottom\" ss:WrapText=\"1\"/>\r\n " +
                        "<Font ss:FontName=\"Arial\" x:Family=\"Swiss\" ss:Size=\"8\" ss:Bold=\"1\"/>\r\n " +
                        "</Style>\r\n " +
                    "</Styles>\r\n ";
            const string endExcelXML = "</Workbook>";

            int colCount = 0;
            int rowCount = 0;
            int sheetCount = 1;
            int i = 1;
            //DataRow[] j;
            /*
            <xml version>
            <Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet"
            xmlns:o="urn:schemas-microsoft-com:office:office"
            xmlns:x="urn:schemas-microsoft-com:office:excel"
            xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet">
            <Styles>
            <Style ss:ID="Default" ss:Name="Normal">
                <Alignment ss:Vertical="Bottom"/>
                <Borders/>
                <Font/>
                <Interior/>
                <NumberFormat/>
                <Protection/>
            </Style>
            <Style ss:ID="BoldColumn">
                <Font x:Family="Swiss" ss:Bold="1"/>
            </Style>
            <Style ss:ID="StringLiteral">
                <NumberFormat ss:Format="@"/>
            </Style>
            <Style ss:ID="Decimal">
                <NumberFormat ss:Format="0.0000"/>
            </Style>
            <Style ss:ID="Integer">
                <NumberFormat ss:Format="0"/>
            </Style>
            <Style ss:ID="DateLiteral">
                <NumberFormat ss:Format="mm/dd/yyyy;@"/>
            </Style>
            </Styles>
            <Worksheet ss:Name="Sheet1">
            </Worksheet>
            </Workbook>
            */
            excelDoc.Write(startExcelXML);
            //excelDoc.Write("<Font ss:FontName=\"Calibri\" x:Family=\"Swiss\" ss:Size=\"8\" ss:Bold=\"1\"/>");
            excelDoc.Write("<Worksheet ss:Name=\"Sheet" + sheetCount + "\">");
            excelDoc.Write("<Table>");

            //excelDoc.Write("<Column ss:Index=\"1\" ss:AutoFitWidth=\"0\" ss:Width=\"50\"/>");
            //excelDoc.Write("<Column ss:Index=\"2\" ss:AutoFitWidth=\"0\" ss:Width=\"50\"/>");
            //excelDoc.Write("<Column ss:Index=\"3\" ss:AutoFitWidth=\"0\" ss:Width=\"40\"/>");
            //excelDoc.Write("<Column ss:Index=\"4\" ss:AutoFitWidth=\"0\" ss:Width=\"150\"/>");
            //excelDoc.Write("<Column ss:Index=\"5\" ss:AutoFitWidth=\"0\" ss:Width=\"40\"/>");
            //excelDoc.Write("<Column ss:Index=\"6\" ss:AutoFitWidth=\"0\" ss:Width=\"40\"/>");

            i = 0;
            foreach (DataRow x in ds.Tables[0].Rows)
            {
                i += 1;
                excelDoc.Write("<Column ss:Index=\"" + i.ToString() + "\" ss:AutoFitWidth=\"0\" ss:Width=\"" + ((string)x["TitLen"]).Trim() + "\"/>");
            }

            excelDoc.Write("<Row ss:Height=\"15.75\">");
            excelDoc.Write("<Cell ss:StyleID=\"BoldColumn2\"><Data ss:Type=\"String\">" + Titolo + "</Data></Cell>");
            excelDoc.Write("</Row>");

            excelDoc.Write("<Row ss:AutoFitHeight=\"0\" ss:Height=\"33.75\" ss:StyleID=\"s67\">");
            foreach (DataRow x in ds.Tables[0].Rows)
            {
                //excelDoc.Write("<Cell ss:StyleID=\"BoldColumn\"><Data ss:Type=\"String\">");
                excelDoc.Write("<Cell><Data ss:Type=\"String\">");
                //excelDoc.Write(source.Tables[0].Columns[x].ColumnName);
                excelDoc.Write(x["TitDes"]);
                excelDoc.Write("</Data></Cell>");
            }
            excelDoc.Write("</Row>");

            foreach (DataRow x in ds.Tables[1].Rows)
            {

                colCount = 0;
                rowCount++;
                // string s2 = "";
                //if the number of rows is > 64000 create a new page to continue output
                if (rowCount == 64000)
                {
                    rowCount = 0;
                    sheetCount++;
                    excelDoc.Write("</Table>");
                    excelDoc.Write(" </Worksheet>");
                    excelDoc.Write("<Worksheet ss:Name=\"Sheet" + sheetCount + "\">");
                    excelDoc.Write("<Table>");
                }
                excelDoc.Write("<Row>");
                //for (int y = 0; y < source.Tables[1].Columns.Count; y++)
                foreach (DataRow y in ds.Tables[0].Rows)
                {

                    string sFld = (string)y["TitFld"];

                    //s2 += source.Tables[1].Columns[y].ColumnName + " | ";
                    //j = source.Tables[0].Select("TitFld='" + source.Tables[1].Columns[y].ColumnName + "'");
                    //if (j.Length > 0)
                    //{
                    colCount++;

                    System.Type rowType;
                    rowType = x[sFld].GetType();
                    switch (rowType.ToString())
                    {
                        case "System.String":
                            string XMLstring = x[sFld].ToString();
                            XMLstring = XMLstring.Trim();
                            XMLstring = XMLstring.Replace("&", "&");
                            XMLstring = XMLstring.Replace(">", ">");
                            XMLstring = XMLstring.Replace("<", "<");
                            if (ds.Tables[1].Columns[sFld].ColumnName == "fat_cli")
                                excelDoc.Write("<Cell ss:StyleID=\"InCentro\"><Data ss:Type=\"String\">");
                            else
                                excelDoc.Write("<Cell ss:StyleID=\"StringLiteral\"><Data ss:Type=\"String\">");
                            excelDoc.Write(XMLstring);
                            excelDoc.Write("</Data></Cell>");
                            break;
                        case "System.DateTime":
                            //Excel has a specific Date Format of YYYY-MM-DD followed by  
                            //the letter 'T' then hh:mm:sss.lll Example 2005-01-31T24:01:21.000
                            //The Following Code puts the date stored in XMLDate 
                            //to the format above
                            DateTime XMLDate = (DateTime)x[sFld];
                            string XMLDatetoString = ""; //Excel Converted Date
                            if (XMLDate.Year.ToString() == "2050")
                                excelDoc.Write("<Cell ss:StyleID=\"StringLiteral\"><Data ss:Type=\"String\"></Data></Cell>");
                            else
                            {
                                XMLDatetoString = XMLDate.Year.ToString() +
                                        "-" +
                                        (XMLDate.Month < 10 ? "0" +
                                        XMLDate.Month.ToString() : XMLDate.Month.ToString()) +
                                        "-" +
                                        (XMLDate.Day < 10 ? "0" +
                                        XMLDate.Day.ToString() : XMLDate.Day.ToString()) +
                                        "T" +
                                        (XMLDate.Hour < 10 ? "0" +
                                        XMLDate.Hour.ToString() : XMLDate.Hour.ToString()) +
                                        ":" +
                                        (XMLDate.Minute < 10 ? "0" +
                                        XMLDate.Minute.ToString() : XMLDate.Minute.ToString()) +
                                        ":" +
                                        (XMLDate.Second < 10 ? "0" +
                                        XMLDate.Second.ToString() : XMLDate.Second.ToString()) +
                                        ".000";
                                excelDoc.Write("<Cell ss:StyleID=\"DateLiteral\"><Data ss:Type=\"DateTime\">");
                                excelDoc.Write(XMLDatetoString);
                                excelDoc.Write("</Data></Cell>");
                            }
                            break;
                        case "System.Boolean":
                            excelDoc.Write("<Cell ss:StyleID=\"StringLiteral\"><Data ss:Type=\"String\">");
                            excelDoc.Write(x[sFld].ToString());
                            excelDoc.Write("</Data></Cell>");
                            break;
                        case "System.Int16":
                        case "System.Int32":
                        case "System.Int64":
                        case "System.Byte":
                            excelDoc.Write("<Cell ss:StyleID=\"Integer\">" +
                                    "<Data ss:Type=\"Number\">");
                            excelDoc.Write(x[sFld].ToString());
                            excelDoc.Write("</Data></Cell>");
                            break;
                        case "System.Decimal":
                        case "System.Double":

                            if (ds.Tables[1].Columns[sFld].ColumnName == "ana_pxc" || ds.Tables[1].Columns[sFld].ColumnName == "ana_con")
                                excelDoc.Write("<Cell ss:StyleID=\"Integer\"><Data ss:Type=\"Number\">");
                            else
                                //excelDoc.Write("<Cell ss:StyleID=\"Decimal\"><Data ss:Type=\"Number\">");
                                excelDoc.Write("<Cell><Data ss:Type=\"Number\">");

                            s = x[sFld].ToString();
                            s = s.Replace(",", ".");
                            excelDoc.Write(s);
                            excelDoc.Write("</Data></Cell>");
                            break;
                        case "System.DBNull":
                            excelDoc.Write("<Cell ss:StyleID=\"StringLiteral\"><Data ss:Type=\"String\">");
                            excelDoc.Write("");
                            excelDoc.Write("</Data></Cell>");
                            break;
                        default:
                            throw (new Exception(rowType.ToString() + " not handled."));
                    }
                    //}
                }

                //if ((string)x["ast_art"] == "42406")
                //    Console.WriteLine("aaa");

                ////for (int i = colCount; i <= 15; i++)
                ////{
                ////excelDoc.Write("<Cell ss:StyleID=\"StringLiteral\"><Data ss:Type=\"String\"></Data></Cell>");
                ////}

                ////s = "0";
                ////s = s.Trim();
                ////s = s.Replace("&", "&");
                ////s = s.Replace(">", ">");
                ////s = s.Replace("<", "<");

                ////excelDoc.Write("<Cell ss:StyleID=\"Decimal\" ss:Formula=\"=G2+H2\"><Data ss:Type=\"Number\">");

                ////=((M2/(1+(I2/100)))-L2)/M2
                ////excelDoc.Write("<Cell ss:StyleID=\"Decimal\" ss:Formula=\"=SUM(RC[-4]:RC[-3])\"><Data ss:Type=\"Number\">0</Data></Cell>");

                //if ((decimal)x["pam_iva"] > 0 && (decimal)x["pam_l14"] > 0)
                //    s = "(((RC[-1]/(1+(RC[-4]/100)))-RC[-2])/RC[-1])*100";
                //else
                //    s = "0";

                //excelDoc.Write("<Cell ss:StyleID=\"Decimal\" ss:Formula=\"=" + s + "\"><Data ss:Type=\"Number\">0</Data></Cell>");

                ////excelDoc.Write("0");
                ////excelDoc.Write("</Data></Cell>");

                //excelDoc.Write("<Cell ss:StyleID=\"StringLiteral\"><Data ss:Type=\"String\">AAA</Data></Cell>");

                excelDoc.Write("</Row>");

            }

            if (Foot != "")
            {
                excelDoc.Write("<Row ss:Height=\"15.75\">");
                excelDoc.Write("<Cell ss:StyleID=\"BoldColumn2\"><Data ss:Type=\"String\">" + Foot + "</Data></Cell>");
                excelDoc.Write("</Row>");
            }

            excelDoc.Write("</Table>");
            excelDoc.Write(" </Worksheet>");
            excelDoc.Write(endExcelXML);
            excelDoc.Close();

            Processo(fileName);

        }

        public void exportToExcel2(DataSet ds, string fileName, string Titolo, string Foot)
        {

            System.IO.StreamWriter excelDoc;
            string s = "";

            excelDoc = new System.IO.StreamWriter(fileName);
            const string startExcelXML =
                    "<xml version>\r\n" +
                    "<Workbook xmlns=\"urn:schemas-microsoft-com:office:spreadsheet\"\r\n " +
                    "xmlns:o=\"urn:schemas-microsoft-com:office:office\"\r\n " +
                    "xmlns:x=\"urn:schemas-microsoft-com:office:excel\"\r\n " +
                    "xmlns:ss=\"urn:schemas-microsoft-com:office:spreadsheet\">\r\n " +
                    "<Styles>\r\n " +
                    "<Style ss:ID=\"Default\" ss:Name=\"Normal\">\r\n " +
                        "<Alignment ss:Vertical=\"Bottom\"/>\r\n <Borders/>\r\n " +
                        "<Font ss:FontName=\"Arial\" x:Family=\"Swiss\" ss:Size=\"8\"/>\r\n " +
                        "<Interior/>\r\n " +
                        "<NumberFormat/>\r\n " +
                        "<Protection/>\r\n " +
                    "</Style>\r\n " +
                    "<Style ss:ID=\"InCentro\">\r\n " +
                        "<Alignment ss:Horizontal=\"Center\" ss:Vertical=\"Bottom\"/>\r\n " +
                    "</Style>\r\n " +
                    "<Style ss:ID=\"BoldColumn\">\r\n " +
                        "<Font ss:FontName=\"Arial\" x:Family=\"Swiss\" ss:Size=\"9\" ss:Bold=\"1\"/>\r\n " +
                    "</Style>\r\n " +
                    "<Style ss:ID=\"BoldColumn2\">\r\n " +
                        "<Font x:Family=\"Swiss\" ss:Bold=\"1\" ss:Size=\"12\"/>\r\n " +
                    "</Style>\r\n " +
                    "<Style ss:ID=\"StringLiteral\">\r\n " +
                        "<NumberFormat ss:Format=\"@\"/>\r\n " +
                    "</Style>\r\n " +
                    "<Style ss:ID=\"Decimal\">\r\n " +
                        "<NumberFormat ss:Format=\"0.00##\"/>\r\n " +
                    "</Style>\r\n " +
                    "<Style ss:ID=\"Integer\">\r\n " +
                        "<NumberFormat ss:Format=\"0\"/>\r\n " +
                    "</Style>\r\n " +
                    "<Style ss:ID=\"DateLiteral\">\r\n " +
                        "<NumberFormat ss:Format=\"dd/mm/yyyy;@\"/>\r\n " +
                    "</Style>\r\n " +
                    "<Style ss:ID=\"s67\">\r\n " +
                        "<Alignment ss:Vertical=\"Bottom\" ss:WrapText=\"1\"/>\r\n " +
                        "<Font ss:FontName=\"Arial\" x:Family=\"Swiss\" ss:Size=\"8\" ss:Bold=\"1\"/>\r\n " +
                        "</Style>\r\n " +
                    "</Styles>\r\n ";
            const string endExcelXML = "</Workbook>";

            int colCount = 0;
            int rowCount = 0;
            int sheetCount = 1;
            int i = 1;
            //DataRow[] j;
            excelDoc.Write(startExcelXML);
            excelDoc.Write("<Worksheet ss:Name=\"Sheet" + sheetCount + "\">");
            excelDoc.Write("<Table>");

            i = 0;
            foreach (DataRow x in ds.Tables[0].Rows)
            {
                i += 1;
                excelDoc.Write("<Column ss:Index=\"" + i.ToString() + "\" ss:AutoFitWidth=\"0\" ss:Width=\"" + ((string)x["TitLen"]).Trim() + "\"/>");
            }

            excelDoc.Write("<Row ss:Height=\"15.75\">");
            excelDoc.Write("<Cell ss:StyleID=\"BoldColumn2\"><Data ss:Type=\"String\">" + Titolo + "</Data></Cell>");
            excelDoc.Write("</Row>");

            excelDoc.Write("<Row ss:AutoFitHeight=\"0\" ss:Height=\"33.75\" ss:StyleID=\"s67\">");
            foreach (DataRow x in ds.Tables[0].Rows)
            {
                excelDoc.Write("<Cell><Data ss:Type=\"String\">");
                excelDoc.Write(x["TitDes"]);
                excelDoc.Write("</Data></Cell>");
            }
            excelDoc.Write("</Row>");

            foreach (DataRow x in ds.Tables[1].Rows)
            {

                colCount = 0;
                rowCount++;
                //string s2 = "";
                //if the number of rows is > 64000 create a new page to continue output
                if (rowCount == 64000)
                {
                    rowCount = 0;
                    sheetCount++;
                    excelDoc.Write("</Table>");
                    excelDoc.Write(" </Worksheet>");
                    excelDoc.Write("<Worksheet ss:Name=\"Sheet" + sheetCount + "\">");
                    excelDoc.Write("<Table>");
                }

                if ((string)x["lav_tip"] == "1")
                    excelDoc.Write("<Row ss:AutoFitHeight=\"0\" ss:Height=\"30.75\" ss:StyleID=\"s67\">");
                else
                    excelDoc.Write("<Row>");

                foreach (DataRow y in ds.Tables[0].Rows)
                {

                    string sFld = (string)y["TitFld"];

                    colCount++;

                    System.Type rowType;
                    rowType = x[sFld].GetType();
                    switch (rowType.ToString())
                    {
                        case "System.String":
                            string XMLstring = x[sFld].ToString();
                            XMLstring = XMLstring.Trim();
                            XMLstring = XMLstring.Replace("&", "&");
                            XMLstring = XMLstring.Replace(">", ">");
                            XMLstring = XMLstring.Replace("<", "<");
                            if ((string)x["lav_tip"] != "3")
                                excelDoc.Write("<Cell ss:StyleID=\"BoldColumn\"><Data ss:Type=\"String\">");
                            else
                                excelDoc.Write("<Cell ss:StyleID=\"StringLiteral\"><Data ss:Type=\"String\">");
                            excelDoc.Write(XMLstring);
                            excelDoc.Write("</Data></Cell>");
                            break;
                        case "System.DateTime":
                            //Excel has a specific Date Format of YYYY-MM-DD followed by  
                            //the letter 'T' then hh:mm:sss.lll Example 2005-01-31T24:01:21.000
                            //The Following Code puts the date stored in XMLDate 
                            //to the format above
                            DateTime XMLDate = (DateTime)x[sFld];
                            string XMLDatetoString = ""; //Excel Converted Date
                            if (XMLDate.Year.ToString() == "2050")
                                excelDoc.Write("<Cell ss:StyleID=\"StringLiteral\"><Data ss:Type=\"String\"></Data></Cell>");
                            else
                            {
                                XMLDatetoString = XMLDate.Year.ToString() +
                                        "-" +
                                        (XMLDate.Month < 10 ? "0" +
                                        XMLDate.Month.ToString() : XMLDate.Month.ToString()) +
                                        "-" +
                                        (XMLDate.Day < 10 ? "0" +
                                        XMLDate.Day.ToString() : XMLDate.Day.ToString()) +
                                        "T" +
                                        (XMLDate.Hour < 10 ? "0" +
                                        XMLDate.Hour.ToString() : XMLDate.Hour.ToString()) +
                                        ":" +
                                        (XMLDate.Minute < 10 ? "0" +
                                        XMLDate.Minute.ToString() : XMLDate.Minute.ToString()) +
                                        ":" +
                                        (XMLDate.Second < 10 ? "0" +
                                        XMLDate.Second.ToString() : XMLDate.Second.ToString()) +
                                        ".000";
                                excelDoc.Write("<Cell ss:StyleID=\"DateLiteral\"><Data ss:Type=\"DateTime\">");
                                excelDoc.Write(XMLDatetoString);
                                excelDoc.Write("</Data></Cell>");
                            }
                            break;
                        case "System.Boolean":
                            excelDoc.Write("<Cell ss:StyleID=\"StringLiteral\"><Data ss:Type=\"String\">");
                            excelDoc.Write(x[sFld].ToString());
                            excelDoc.Write("</Data></Cell>");
                            break;
                        case "System.Int16":
                        case "System.Int32":
                        case "System.Int64":
                        case "System.Byte":
                            excelDoc.Write("<Cell ss:StyleID=\"Integer\">" +
                                    "<Data ss:Type=\"Number\">");
                            excelDoc.Write(x[sFld].ToString());
                            excelDoc.Write("</Data></Cell>");
                            break;
                        case "System.Decimal":
                        case "System.Double":

                            if ((string)x["lav_tip"] != "3")
                                excelDoc.Write("<Cell ss:StyleID=\"BoldColumn\"><Data ss:Type=\"String\">");
                            else
                                excelDoc.Write("<Cell ss:StyleID=\"Decimal\"><Data ss:Type=\"Number\">");
                            s = x[sFld].ToString();
                            s = s.Replace(",", ".");
                            excelDoc.Write(s);
                            excelDoc.Write("</Data></Cell>");
                            break;
                        case "System.DBNull":
                            excelDoc.Write("<Cell ss:StyleID=\"StringLiteral\"><Data ss:Type=\"String\">");
                            excelDoc.Write("");
                            excelDoc.Write("</Data></Cell>");
                            break;
                        default:
                            throw (new Exception(rowType.ToString() + " not handled."));
                    }
                    //}
                }

                excelDoc.Write("</Row>");

            }

            if (Foot != "")
            {
                excelDoc.Write("<Row ss:Height=\"15.75\">");
                excelDoc.Write("<Cell ss:StyleID=\"BoldColumn2\"><Data ss:Type=\"String\">" + Foot + "</Data></Cell>");
                excelDoc.Write("</Row>");
            }

            excelDoc.Write("</Table>");
            excelDoc.Write(" </Worksheet>");
            excelDoc.Write(endExcelXML);
            excelDoc.Close();

            Processo(fileName);

        }

        public void exportToExcel3(DataSet ds, string fileName, string Titolo, string Foot)
        {

            System.IO.StreamWriter excelDoc;
            string s = "";

            excelDoc = new System.IO.StreamWriter(fileName);
            const string startExcelXML =
                    "<xml version>\r\n" +
                    "<Workbook xmlns=\"urn:schemas-microsoft-com:office:spreadsheet\"\r\n " +
                    "xmlns:o=\"urn:schemas-microsoft-com:office:office\"\r\n " +
                    "xmlns:x=\"urn:schemas-microsoft-com:office:excel\"\r\n " +
                    "xmlns:ss=\"urn:schemas-microsoft-com:office:spreadsheet\">\r\n " +
                    "<Styles>\r\n " +
                    "<Style ss:ID=\"Default\" ss:Name=\"Normal\">\r\n " +
                        "<Alignment ss:Vertical=\"Bottom\"/>\r\n <Borders/>\r\n " +
                        "<Font ss:FontName=\"Arial\" x:Family=\"Swiss\" ss:Size=\"8\"/>\r\n " +
                        "<Interior/>\r\n " +
                        "<NumberFormat/>\r\n " +
                        "<Protection/>\r\n " +
                    "</Style>\r\n " +
                    "<Style ss:ID=\"InCentro\">\r\n " +
                        "<Alignment ss:Horizontal=\"Center\" ss:Vertical=\"Bottom\"/>\r\n " +
                    "</Style>\r\n " +
                    "<Style ss:ID=\"BoldColumn\">\r\n " +
                        "<Font ss:FontName=\"Arial\" x:Family=\"Swiss\" ss:Size=\"9\" ss:Bold=\"1\"/>\r\n " +
                    "</Style>\r\n " +
                    "<Style ss:ID=\"BoldColumn2\">\r\n " +
                        "<Font x:Family=\"Swiss\" ss:Bold=\"1\" ss:Size=\"12\"/>\r\n " +
                    "</Style>\r\n " +
                    "<Style ss:ID=\"StringLiteral\">\r\n " +
                        "<NumberFormat ss:Format=\"@\"/>\r\n " +
                    "</Style>\r\n " +
                    "<Style ss:ID=\"Decimal\">\r\n " +
                        "<NumberFormat ss:Format=\"0.00##\"/>\r\n " +
                    "</Style>\r\n " +
                    "<Style ss:ID=\"Integer\">\r\n " +
                        "<NumberFormat ss:Format=\"0\"/>\r\n " +
                    "</Style>\r\n " +
                    "<Style ss:ID=\"DateLiteral\">\r\n " +
                        "<NumberFormat ss:Format=\"dd/mm/yyyy;@\"/>\r\n " +
                    "</Style>\r\n " +
                    "<Style ss:ID=\"s67\">\r\n " +
                        "<Alignment ss:Vertical=\"Bottom\" ss:WrapText=\"1\"/>\r\n " +
                        "<Font ss:FontName=\"Arial\" x:Family=\"Swiss\" ss:Size=\"8\" ss:Bold=\"1\"/>\r\n " +
                        "</Style>\r\n " +
                    "</Styles>\r\n ";
            const string endExcelXML = "</Workbook>";

            int colCount = 0;
            int rowCount = 0;
            int sheetCount = 1;
            int i = 1;
            //DataRow[] j;
            excelDoc.Write(startExcelXML);
            excelDoc.Write("<Worksheet ss:Name=\"Sheet" + sheetCount + "\">");
            excelDoc.Write("<Table>");

            i = 0;
            foreach (DataRow x in ds.Tables[0].Rows)
            {
                i += 1;
                excelDoc.Write("<Column ss:Index=\"" + i.ToString() + "\" ss:AutoFitWidth=\"0\" ss:Width=\"" + ((string)x["TitLen"]).Trim() + "\"/>");
            }

            excelDoc.Write("<Row ss:Height=\"15.75\">");
            excelDoc.Write("<Cell ss:StyleID=\"BoldColumn2\"><Data ss:Type=\"String\">" + Titolo + "</Data></Cell>");
            excelDoc.Write("</Row>");

            excelDoc.Write("<Row ss:AutoFitHeight=\"0\" ss:Height=\"33.75\" ss:StyleID=\"s67\">");
            foreach (DataRow x in ds.Tables[0].Rows)
            {
                excelDoc.Write("<Cell><Data ss:Type=\"String\">");
                excelDoc.Write(x["TitDes"]);
                excelDoc.Write("</Data></Cell>");
            }
            excelDoc.Write("</Row>");

            foreach (DataRow x in ds.Tables[1].Rows)
            {

                colCount = 0;
                rowCount++;
                //string s2 = "";
                //if the number of rows is > 64000 create a new page to continue output
                if (rowCount == 64000)
                {
                    rowCount = 0;
                    sheetCount++;
                    excelDoc.Write("</Table>");
                    excelDoc.Write(" </Worksheet>");
                    excelDoc.Write("<Worksheet ss:Name=\"Sheet" + sheetCount + "\">");
                    excelDoc.Write("<Table>");
                }

                if ((string)x["mov_tip"] == "1")
                    excelDoc.Write("<Row ss:AutoFitHeight=\"0\" ss:Height=\"30.75\" ss:StyleID=\"s67\">");
                else
                    excelDoc.Write("<Row>");

                foreach (DataRow y in ds.Tables[0].Rows)
                {

                    string sFld = (string)y["TitFld"];

                    colCount++;

                    System.Type rowType;
                    rowType = x[sFld].GetType();
                    switch (rowType.ToString())
                    {
                        case "System.String":
                            string XMLstring = x[sFld].ToString();
                            XMLstring = XMLstring.Trim();
                            XMLstring = XMLstring.Replace("&", "&");
                            XMLstring = XMLstring.Replace(">", ">");
                            XMLstring = XMLstring.Replace("<", "<");
                            if ((string)x["mov_tip"] == "1" || (string)x["mov_tip"] == "4")
                                excelDoc.Write("<Cell ss:StyleID=\"BoldColumn\"><Data ss:Type=\"String\">");
                            else
                                excelDoc.Write("<Cell ss:StyleID=\"StringLiteral\"><Data ss:Type=\"String\">");
                            excelDoc.Write(XMLstring);
                            excelDoc.Write("</Data></Cell>");
                            break;
                        case "System.DateTime":
                            //Excel has a specific Date Format of YYYY-MM-DD followed by  
                            //the letter 'T' then hh:mm:sss.lll Example 2005-01-31T24:01:21.000
                            //The Following Code puts the date stored in XMLDate 
                            //to the format above
                            DateTime XMLDate = (DateTime)x[sFld];
                            string XMLDatetoString = ""; //Excel Converted Date
                            if (XMLDate.Year.ToString() == "2050")
                                excelDoc.Write("<Cell ss:StyleID=\"StringLiteral\"><Data ss:Type=\"String\"></Data></Cell>");
                            else
                            {
                                XMLDatetoString = XMLDate.Year.ToString() +
                                        "-" +
                                        (XMLDate.Month < 10 ? "0" +
                                        XMLDate.Month.ToString() : XMLDate.Month.ToString()) +
                                        "-" +
                                        (XMLDate.Day < 10 ? "0" +
                                        XMLDate.Day.ToString() : XMLDate.Day.ToString()) +
                                        "T" +
                                        (XMLDate.Hour < 10 ? "0" +
                                        XMLDate.Hour.ToString() : XMLDate.Hour.ToString()) +
                                        ":" +
                                        (XMLDate.Minute < 10 ? "0" +
                                        XMLDate.Minute.ToString() : XMLDate.Minute.ToString()) +
                                        ":" +
                                        (XMLDate.Second < 10 ? "0" +
                                        XMLDate.Second.ToString() : XMLDate.Second.ToString()) +
                                        ".000";
                                excelDoc.Write("<Cell ss:StyleID=\"DateLiteral\"><Data ss:Type=\"DateTime\">");
                                excelDoc.Write(XMLDatetoString);
                                excelDoc.Write("</Data></Cell>");
                            }
                            break;
                        case "System.Boolean":
                            excelDoc.Write("<Cell ss:StyleID=\"StringLiteral\"><Data ss:Type=\"String\">");
                            excelDoc.Write(x[sFld].ToString());
                            excelDoc.Write("</Data></Cell>");
                            break;
                        case "System.Int16":
                        case "System.Int32":
                        case "System.Int64":
                        case "System.Byte":
                            excelDoc.Write("<Cell ss:StyleID=\"Integer\">" +
                                    "<Data ss:Type=\"Number\">");
                            excelDoc.Write(x[sFld].ToString());
                            excelDoc.Write("</Data></Cell>");
                            break;
                        case "System.Decimal":
                        case "System.Double":

                            if ((string)x["mov_tip"] == "1" || (string)x["mov_tip"] == "4")
                                excelDoc.Write("<Cell ss:StyleID=\"BoldColumn\"><Data ss:Type=\"Number\">");
                            else
                                excelDoc.Write("<Cell ss:StyleID=\"Decimal\"><Data ss:Type=\"Number\">");
                            s = x[sFld].ToString();
                            s = s.Replace(",", ".");
                            excelDoc.Write(s);
                            excelDoc.Write("</Data></Cell>");
                            break;
                        case "System.DBNull":
                            excelDoc.Write("<Cell ss:StyleID=\"StringLiteral\"><Data ss:Type=\"String\">");
                            excelDoc.Write("");
                            excelDoc.Write("</Data></Cell>");
                            break;
                        default:
                            throw (new Exception(rowType.ToString() + " not handled."));
                    }
                    //}
                }

                excelDoc.Write("</Row>");

            }

            if (Foot != "")
            {
                excelDoc.Write("<Row ss:Height=\"15.75\">");
                excelDoc.Write("<Cell ss:StyleID=\"BoldColumn2\"><Data ss:Type=\"String\">" + Foot + "</Data></Cell>");
                excelDoc.Write("</Row>");
            }

            excelDoc.Write("</Table>");
            excelDoc.Write(" </Worksheet>");
            excelDoc.Write(endExcelXML);
            excelDoc.Close();

            Processo(fileName);

        }

        public void exportToExcel4(DataSet ds, string fileName, string Titolo, string Foot)
        {

            System.IO.StreamWriter excelDoc;
            string s = "";

            excelDoc = new System.IO.StreamWriter(fileName);
            const string startExcelXML =
                    "<xml version>\r\n" +
                    "<Workbook xmlns=\"urn:schemas-microsoft-com:office:spreadsheet\"\r\n " +
                    "xmlns:o=\"urn:schemas-microsoft-com:office:office\"\r\n " +
                    "xmlns:x=\"urn:schemas-microsoft-com:office:excel\"\r\n " +
                    "xmlns:ss=\"urn:schemas-microsoft-com:office:spreadsheet\">\r\n " +
                    "<Styles>\r\n " +
                    "<Style ss:ID=\"Default\" ss:Name=\"Normal\">\r\n " +
                        "<Alignment ss:Vertical=\"Bottom\"/>\r\n <Borders/>\r\n " +
                        "<Font ss:FontName=\"Arial\" x:Family=\"Swiss\" ss:Size=\"8\"/>\r\n " +
                        "<Interior/>\r\n " +
                        "<NumberFormat/>\r\n " +
                        "<Protection/>\r\n " +
                    "</Style>\r\n " +
                    "<Style ss:ID=\"InCentro\">\r\n " +
                        "<Alignment ss:Horizontal=\"Center\" ss:Vertical=\"Bottom\"/>\r\n " +
                    "</Style>\r\n " +
                    "<Style ss:ID=\"BoldColumn\">\r\n " +
                        "<Font ss:FontName=\"Arial\" x:Family=\"Swiss\" ss:Size=\"9\" ss:Bold=\"1\"/>\r\n " +
                    "</Style>\r\n " +
                    "<Style ss:ID=\"BoldColumn2\">\r\n " +
                        "<Font x:Family=\"Swiss\" ss:Bold=\"1\" ss:Size=\"12\"/>\r\n " +
                    "</Style>\r\n " +
                    "<Style ss:ID=\"StringLiteral\">\r\n " +
                        "<NumberFormat ss:Format=\"@\"/>\r\n " +
                    "</Style>\r\n " +
                    "<Style ss:ID=\"Decimal\">\r\n " +
                        "<NumberFormat ss:Format=\"0.00##\"/>\r\n " +
                    "</Style>\r\n " +
                    "<Style ss:ID=\"Integer\">\r\n " +
                        "<NumberFormat ss:Format=\"0\"/>\r\n " +
                    "</Style>\r\n " +
                    "<Style ss:ID=\"DateLiteral\">\r\n " +
                        "<NumberFormat ss:Format=\"dd/mm/yyyy;@\"/>\r\n " +
                    "</Style>\r\n " +
                    "<Style ss:ID=\"s67\">\r\n " +
                        "<Alignment ss:Vertical=\"Bottom\" ss:WrapText=\"1\"/>\r\n " +
                        "<Font ss:FontName=\"Arial\" x:Family=\"Swiss\" ss:Size=\"8\" ss:Bold=\"1\"/>\r\n " +
                        "</Style>\r\n " +
                    "</Styles>\r\n ";
            const string endExcelXML = "</Workbook>";

            int colCount = 0;
            int rowCount = 0;
            int sheetCount = 1;
            int i = 1;
            //DataRow[] j;
            excelDoc.Write(startExcelXML);
            excelDoc.Write("<Worksheet ss:Name=\"Sheet" + sheetCount + "\">");
            excelDoc.Write("<Table>");

            i = 0;
            foreach (DataRow x in ds.Tables[0].Rows)
            {
                i += 1;
                excelDoc.Write("<Column ss:Index=\"" + i.ToString() + "\" ss:AutoFitWidth=\"0\" ss:Width=\"" + ((string)x["TitLen"]).Trim() + "\"/>");
            }

            excelDoc.Write("<Row ss:Height=\"15.75\">");
            excelDoc.Write("<Cell ss:StyleID=\"BoldColumn2\"><Data ss:Type=\"String\">" + Titolo + "</Data></Cell>");
            excelDoc.Write("</Row>");

            excelDoc.Write("<Row ss:AutoFitHeight=\"0\" ss:Height=\"33.75\" ss:StyleID=\"s67\">");
            foreach (DataRow x in ds.Tables[0].Rows)
            {
                excelDoc.Write("<Cell><Data ss:Type=\"String\">");
                excelDoc.Write(x["TitDes"]);
                excelDoc.Write("</Data></Cell>");
            }
            excelDoc.Write("</Row>");

            foreach (DataRow x in ds.Tables[1].Rows)
            {

                colCount = 0;
                rowCount++;
                //string s2 = "";
                //if the number of rows is > 64000 create a new page to continue output
                if (rowCount == 64000)
                {
                    rowCount = 0;
                    sheetCount++;
                    excelDoc.Write("</Table>");
                    excelDoc.Write(" </Worksheet>");
                    excelDoc.Write("<Worksheet ss:Name=\"Sheet" + sheetCount + "\">");
                    excelDoc.Write("<Table>");
                }

                if ((string)x["mov_tip"] == "1")
                    excelDoc.Write("<Row ss:AutoFitHeight=\"0\" ss:Height=\"30.75\" ss:StyleID=\"s67\">");
                else
                    excelDoc.Write("<Row>");

                foreach (DataRow y in ds.Tables[0].Rows)
                {

                    string sFld = (string)y["TitFld"];

                    colCount++;

                    System.Type rowType;
                    rowType = x[sFld].GetType();
                    switch (rowType.ToString())
                    {
                        case "System.String":
                            string XMLstring = x[sFld].ToString();
                            XMLstring = XMLstring.Trim();
                            XMLstring = XMLstring.Replace("&", "&");
                            XMLstring = XMLstring.Replace(">", ">");
                            XMLstring = XMLstring.Replace("<", "<");
                            if ((string)x["mov_tip"] == "1")
                                excelDoc.Write("<Cell ss:StyleID=\"BoldColumn2\"><Data ss:Type=\"String\">");
                            else if ((string)x["mov_tip"] == "3" || (string)x["mov_tip"] == "4" || (string)x["mov_tip"] == "5" || (string)x["mov_tip"] == "8")
                                excelDoc.Write("<Cell ss:StyleID=\"BoldColumn\"><Data ss:Type=\"String\">");
                            else
                                excelDoc.Write("<Cell ss:StyleID=\"StringLiteral\"><Data ss:Type=\"String\">");
                            excelDoc.Write(XMLstring);
                            excelDoc.Write("</Data></Cell>");
                            break;
                        case "System.DateTime":
                            //Excel has a specific Date Format of YYYY-MM-DD followed by  
                            //the letter 'T' then hh:mm:sss.lll Example 2005-01-31T24:01:21.000
                            //The Following Code puts the date stored in XMLDate 
                            //to the format above
                            DateTime XMLDate = (DateTime)x[sFld];
                            string XMLDatetoString = ""; //Excel Converted Date
                            if (XMLDate.Year.ToString() == "2050")
                                excelDoc.Write("<Cell ss:StyleID=\"StringLiteral\"><Data ss:Type=\"String\"></Data></Cell>");
                            else
                            {
                                XMLDatetoString = XMLDate.Year.ToString() +
                                        "-" +
                                        (XMLDate.Month < 10 ? "0" +
                                        XMLDate.Month.ToString() : XMLDate.Month.ToString()) +
                                        "-" +
                                        (XMLDate.Day < 10 ? "0" +
                                        XMLDate.Day.ToString() : XMLDate.Day.ToString()) +
                                        "T" +
                                        (XMLDate.Hour < 10 ? "0" +
                                        XMLDate.Hour.ToString() : XMLDate.Hour.ToString()) +
                                        ":" +
                                        (XMLDate.Minute < 10 ? "0" +
                                        XMLDate.Minute.ToString() : XMLDate.Minute.ToString()) +
                                        ":" +
                                        (XMLDate.Second < 10 ? "0" +
                                        XMLDate.Second.ToString() : XMLDate.Second.ToString()) +
                                        ".000";
                                excelDoc.Write("<Cell ss:StyleID=\"DateLiteral\"><Data ss:Type=\"DateTime\">");
                                excelDoc.Write(XMLDatetoString);
                                excelDoc.Write("</Data></Cell>");
                            }
                            break;
                        case "System.Boolean":
                            excelDoc.Write("<Cell ss:StyleID=\"StringLiteral\"><Data ss:Type=\"String\">");
                            excelDoc.Write(x[sFld].ToString());
                            excelDoc.Write("</Data></Cell>");
                            break;
                        case "System.Int16":
                        case "System.Int32":
                        case "System.Int64":
                        case "System.Byte":
                            excelDoc.Write("<Cell ss:StyleID=\"Integer\">" +
                                    "<Data ss:Type=\"Number\">");
                            excelDoc.Write(x[sFld].ToString());
                            excelDoc.Write("</Data></Cell>");
                            break;
                        case "System.Decimal":
                        case "System.Double":

                            if ((string)x["mov_tip"] == "1" || (string)x["mov_tip"] == "4")
                                excelDoc.Write("<Cell ss:StyleID=\"BoldColumn\"><Data ss:Type=\"Number\">");
                            else
                                excelDoc.Write("<Cell ss:StyleID=\"Decimal\"><Data ss:Type=\"Number\">");
                            s = x[sFld].ToString();
                            s = s.Replace(",", ".");
                            excelDoc.Write(s);
                            excelDoc.Write("</Data></Cell>");
                            break;
                        case "System.DBNull":
                            excelDoc.Write("<Cell ss:StyleID=\"StringLiteral\"><Data ss:Type=\"String\">");
                            excelDoc.Write("");
                            excelDoc.Write("</Data></Cell>");
                            break;
                        default:
                            throw (new Exception(rowType.ToString() + " not handled."));
                    }
                    //}
                }

                excelDoc.Write("</Row>");

            }

            if (Foot != "")
            {
                excelDoc.Write("<Row ss:Height=\"15.75\">");
                excelDoc.Write("<Cell ss:StyleID=\"BoldColumn2\"><Data ss:Type=\"String\">" + Foot + "</Data></Cell>");
                excelDoc.Write("</Row>");
            }

            excelDoc.Write("</Table>");
            excelDoc.Write(" </Worksheet>");
            excelDoc.Write(endExcelXML);
            excelDoc.Close();

            Processo(fileName);

        }

        public void exportToExcel4Old(DataSet ds, string fileName, string Titolo, string Foot)
        {

            System.IO.StreamWriter excelDoc;
            string s = "";

            excelDoc = new System.IO.StreamWriter(fileName);
            const string startExcelXML =
                    "<xml version>\r\n" +
                    "<Workbook xmlns=\"urn:schemas-microsoft-com:office:spreadsheet\"\r\n " +
                    "xmlns:o=\"urn:schemas-microsoft-com:office:office\"\r\n " +
                    "xmlns:x=\"urn:schemas-microsoft-com:office:excel\"\r\n " +
                    "xmlns:ss=\"urn:schemas-microsoft-com:office:spreadsheet\">\r\n " +
                    "<Styles>\r\n " +
                    "<Style ss:ID=\"Default\" ss:Name=\"Normal\">\r\n " +
                        "<Alignment ss:Vertical=\"Bottom\"/>\r\n <Borders/>\r\n " +
                        "<Font ss:FontName=\"Arial\" x:Family=\"Swiss\" ss:Size=\"8\"/>\r\n " +
                        "<Interior/>\r\n " +
                        "<NumberFormat/>\r\n " +
                        "<Protection/>\r\n " +
                    "</Style>\r\n " +
                    "<Style ss:ID=\"InCentro\">\r\n " +
                        "<Alignment ss:Horizontal=\"Center\" ss:Vertical=\"Bottom\"/>\r\n " +
                    "</Style>\r\n " +
                    "<Style ss:ID=\"BoldColumn0\">\r\n " +
                        "<Font ss:FontName=\"Arial\" x:Family=\"Swiss\" ss:Size=\"9\" ss:Bold=\"1\"/>\r\n " +
                    "</Style>\r\n " +
                    "<Style ss:ID=\"BoldColumn1\">\r\n " +
                        "<Font ss:FontName=\"Arial=\" x:Family=\"Swiss\" ss:Size=\"8\" ss:Bold=\"1\"/>\r\n " +
                    "</Style>\r\n " +
                    "<Style ss:ID=\"BoldColumn2\">\r\n " +
                        "<Font x:Family=\"Swiss\" ss:Bold=\"1\" ss:Size=\"12\"/>\r\n " +
                    "</Style>\r\n " +
                    "<Style ss:ID=\"StringLiteral\">\r\n " +
                        "<NumberFormat ss:Format=\"@\"/>\r\n " +
                    "</Style>\r\n " +
                    "<Style ss:ID=\"Decimal\">\r\n " +
                        "<NumberFormat ss:Format=\"0.00##\"/>\r\n " +
                    "</Style>\r\n " +
                    "<Style ss:ID=\"Integer\">\r\n " +
                        "<NumberFormat ss:Format=\"0\"/>\r\n " +
                    "</Style>\r\n " +
                    "<Style ss:ID=\"DateLiteral\">\r\n " +
                        "<NumberFormat ss:Format=\"dd/mm/yyyy;@\"/>\r\n " +
                    "</Style>\r\n " +
                    "<Style ss:ID=\"s67\">\r\n " +
                        "<Alignment ss:Vertical=\"Bottom\" ss:WrapText=\"1\"/>\r\n " +
                        "<Font ss:FontName=\"Arial\" x:Family=\"Swiss\" ss:Size=\"8\" ss:Bold=\"1\"/>\r\n " +
                        "</Style>\r\n " +
                    "</Styles>\r\n ";
            const string endExcelXML = "</Workbook>";

            int colCount = 0;
            int rowCount = 0;
            int sheetCount = 1;
            int i = 1;
            //DataRow[] j;
            excelDoc.Write(startExcelXML);
            excelDoc.Write("<Worksheet ss:Name=\"Sheet" + sheetCount + "\">");
            excelDoc.Write("<Table>");

            i = 0;
            foreach (DataRow x in ds.Tables[0].Rows)
            {
                i += 1;
                excelDoc.Write("<Column ss:Index=\"" + i.ToString() + "\" ss:AutoFitWidth=\"0\" ss:Width=\"" + ((string)x["TitLen"]).Trim() + "\"/>");
            }

            excelDoc.Write("<Row ss:Height=\"15.75\">");
            excelDoc.Write("<Cell ss:StyleID=\"BoldColumn2\"><Data ss:Type=\"String\">" + Titolo + "</Data></Cell>");
            excelDoc.Write("</Row>");

            excelDoc.Write("<Row ss:AutoFitHeight=\"0\" ss:Height=\"33.75\" ss:StyleID=\"s67\">");
            foreach (DataRow x in ds.Tables[0].Rows)
            {
                excelDoc.Write("<Cell><Data ss:Type=\"String\">");
                excelDoc.Write(x["TitDes"]);
                excelDoc.Write("</Data></Cell>");
            }
            excelDoc.Write("</Row>");

            foreach (DataRow x in ds.Tables[1].Rows)
            {

                colCount = 0;
                rowCount++;
                //string s2 = "";
                //if the number of rows is > 64000 create a new page to continue output
                if (rowCount == 64000)
                {
                    rowCount = 0;
                    sheetCount++;
                    excelDoc.Write("</Table>");
                    excelDoc.Write(" </Worksheet>");
                    excelDoc.Write("<Worksheet ss:Name=\"Sheet" + sheetCount + "\">");
                    excelDoc.Write("<Table>");
                }

                if ((string)x["mov_tip"] == "1")
                    excelDoc.Write("<Row ss:AutoFitHeight=\"0\" ss:Height=\"30.75\" ss:StyleID=\"s67\">");
                else
                    excelDoc.Write("<Row>");

                foreach (DataRow y in ds.Tables[0].Rows)
                {

                    string sFld = (string)y["TitFld"];

                    colCount++;

                    System.Type rowType;
                    rowType = x[sFld].GetType();
                    switch (rowType.ToString())
                    {
                        case "System.String":
                            string XMLstring = x[sFld].ToString();
                            XMLstring = XMLstring.Trim();
                            XMLstring = XMLstring.Replace("&", "&");
                            XMLstring = XMLstring.Replace(">", ">");
                            XMLstring = XMLstring.Replace("<", "<");
                            if ((string)x["mov_tip"] == "1")
                                excelDoc.Write("<Cell ss:StyleID=\"BoldColumn1\"><Data ss:Type=\"Number\">");
                            else if ((string)x["mov_tip"] == "4")
                                excelDoc.Write("<Cell ss:StyleID=\"BoldColumn1\"><Data ss:Type=\"Number\">");
                            else
                                excelDoc.Write("<Cell ss:StyleID=\"StringLiteral\"><Data ss:Type=\"String\">");
                            excelDoc.Write(XMLstring);
                            excelDoc.Write("</Data></Cell>");
                            break;
                        case "System.DateTime":
                            //Excel has a specific Date Format of YYYY-MM-DD followed by  
                            //the letter 'T' then hh:mm:sss.lll Example 2005-01-31T24:01:21.000
                            //The Following Code puts the date stored in XMLDate 
                            //to the format above
                            DateTime XMLDate = (DateTime)x[sFld];
                            string XMLDatetoString = ""; //Excel Converted Date
                            if (XMLDate.Year.ToString() == "2050")
                                excelDoc.Write("<Cell ss:StyleID=\"StringLiteral\"><Data ss:Type=\"String\"></Data></Cell>");
                            else
                            {
                                XMLDatetoString = XMLDate.Year.ToString() +
                                        "-" +
                                        (XMLDate.Month < 10 ? "0" +
                                        XMLDate.Month.ToString() : XMLDate.Month.ToString()) +
                                        "-" +
                                        (XMLDate.Day < 10 ? "0" +
                                        XMLDate.Day.ToString() : XMLDate.Day.ToString()) +
                                        "T" +
                                        (XMLDate.Hour < 10 ? "0" +
                                        XMLDate.Hour.ToString() : XMLDate.Hour.ToString()) +
                                        ":" +
                                        (XMLDate.Minute < 10 ? "0" +
                                        XMLDate.Minute.ToString() : XMLDate.Minute.ToString()) +
                                        ":" +
                                        (XMLDate.Second < 10 ? "0" +
                                        XMLDate.Second.ToString() : XMLDate.Second.ToString()) +
                                        ".000";
                                excelDoc.Write("<Cell ss:StyleID=\"DateLiteral\"><Data ss:Type=\"DateTime\">");
                                excelDoc.Write(XMLDatetoString);
                                excelDoc.Write("</Data></Cell>");
                            }
                            break;
                        case "System.Boolean":
                            excelDoc.Write("<Cell ss:StyleID=\"StringLiteral\"><Data ss:Type=\"String\">");
                            excelDoc.Write(x[sFld].ToString());
                            excelDoc.Write("</Data></Cell>");
                            break;
                        case "System.Int16":
                        case "System.Int32":
                        case "System.Int64":
                        case "System.Byte":
                            excelDoc.Write("<Cell ss:StyleID=\"Integer\">" +
                                    "<Data ss:Type=\"Number\">");
                            excelDoc.Write(x[sFld].ToString());
                            excelDoc.Write("</Data></Cell>");
                            break;
                        case "System.Decimal":
                        case "System.Double":
                            if ((string)x["mov_tip"] == "1")
                                excelDoc.Write("<Cell ss:StyleID=\"BoldColumn1\"><Data ss:Type=\"Number\">");
                            else if ((string)x["mov_tip"] == "4")
                                excelDoc.Write("<Cell ss:StyleID=\"BoldColumn1\"><Data ss:Type=\"Number\">");
                            else
                                excelDoc.Write("<Cell ss:StyleID=\"Decimal\"><Data ss:Type=\"Number\">");
                            s = x[sFld].ToString();
                            s = s.Replace(",", ".");
                            excelDoc.Write(s);
                            excelDoc.Write("</Data></Cell>");
                            break;
                        case "System.DBNull":
                            excelDoc.Write("<Cell ss:StyleID=\"StringLiteral\"><Data ss:Type=\"String\">");
                            excelDoc.Write("");
                            excelDoc.Write("</Data></Cell>");
                            break;
                        default:
                            throw (new Exception(rowType.ToString() + " not handled."));
                    }
                    //}
                }

                excelDoc.Write("</Row>");

            }

            if (Foot != "")
            {
                excelDoc.Write("<Row ss:Height=\"15.75\">");
                excelDoc.Write("<Cell ss:StyleID=\"BoldColumn2\"><Data ss:Type=\"String\">" + Foot + "</Data></Cell>");
                excelDoc.Write("</Row>");
            }

            excelDoc.Write("</Table>");
            excelDoc.Write(" </Worksheet>");
            excelDoc.Write(endExcelXML);
            excelDoc.Close();

            Processo(fileName);
        }

        public void exportToCsv1(string strFld, DataTable tabTab, string strFil, string strTit, string strFoo)
        {
            string s = "";
            string sSep = ";";
            string sRig = "";

            string sFil = strFil;

            for (int i = 0; i <= 100; i++)
            {
                sFil = _strPath + DateTime.Now.ToString("yyyyMMddmmss") + i.ToString() + "_" + strFil + ".csv";
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

            StreamWriter sw = new StreamWriter(sFil, true);
            string[] aFld = strFld.Split(';');

            if (strTit.Trim() != "")
                sw.Write(strTit + _clsDef.CRLF);

            sRig = "";
            foreach (string ss in aFld)
            {
                if (ss != "")
                {
                    string[] a = ss.Split(',');
                    sRig += a[1] + sSep;
                }
            }

            sw.Write(sRig + _clsDef.CRLF);

            foreach (DataRow y in tabTab.Rows)
            {
                sRig = "";

                foreach (string ss in aFld)
                {
                    foreach (DataColumn c in tabTab.Columns)
                    {

                        s = c.ColumnName;

                        if (s == "vep_imp")
                            Console.WriteLine("xxxxxxxxxxx");

                        if (ss.Contains(c.ColumnName))
                        {
                            s = Convert.ToString(y[c]);
                            if (c.DataType.Name == "Boolean")
                            {
                                s = "";
                                if (s == "True")
                                    s = "S";
                            }
                            else if (c.DataType.Name == "DateTime")
                                s = ((DateTime)y[c]).ToShortDateString();
                            else if (c.DataType.Name == "Decimal")
                            {
                                if (DBNull.Value.Equals(y[c]))
                                    s = "";
                                else
                                {
                                    s = ((Decimal)y[c]).ToString();
                                    s = s.Replace(",", ".");
                                }
                            }
                            else if (_clsFun.Numerico(s))
                            {
                                //if (Convert.ToDouble(s) > 999999)
                                    s = "'" + s.Trim();
                            }
                            sRig += s + sSep;
                        }
                    }
                }
                sw.Write(sRig + _clsDef.CRLF);
            }

            if (strFoo != "")
            {
                strFoo = strFoo.Replace(",", sSep);
                sw.Write(strFoo + _clsDef.CRLF);
            }
            ((TextWriter)sw).Flush();
            sw.Close();
            sw.Dispose();

            Processo(sFil);
        }

    }
}
