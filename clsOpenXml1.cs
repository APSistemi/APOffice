using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using System.Windows.Forms;
using System.Diagnostics;
using System.IO;

using DocumentFormat.OpenXml.Drawing.Spreadsheet;
using Xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;
using A = DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Drawing.Pictures;

namespace APOffice
{
    class clsOpenXml1
    {
        clsDefine _clsDef = new clsDefine();
        clsFuncs _clsFun = new clsFuncs();
        clsQuery _clsQry = new clsQuery();

        public clsOpenXml1()
        { }

        public void OLDxlsLibroIngredienti(DataTable tabRow, string strFil)
        {
            string sMsg = "";
            string _strCols = "1, 1, 80;";

            Boolean b = true;

            if(File.Exists(strFil))
            {
                try
                {
                    File.Delete(strFil);
                }
                catch (Exception ex)
                {
                    MessageBox.Show("File excel aperto!", "CONTROLLO FILE", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    b = false;
                }
            }

            if(b)
            {
                for (int ii = 0; ii <= tabRow.Rows.Count - 1; ii++)
                {
                    string s = "";
                    string sFil = strFil;

                    string sDesArt = "";

                    int i = 0;

                    using (SpreadsheetDocument document = SpreadsheetDocument.Create(sFil, SpreadsheetDocumentType.Workbook))
                    {
                        WorkbookPart workbookPart = document.AddWorkbookPart();
                        workbookPart.Workbook = new Workbook();

                        WorksheetPart worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
                        worksheetPart.Worksheet = new Worksheet();

                        // Adding style
                        WorkbookStylesPart stylePart = workbookPart.AddNewPart<WorkbookStylesPart>();
                        stylePart.Stylesheet = GenerateStylesheet();
                        stylePart.Stylesheet.Save();

                        // Setting up columns IMPOSTAZIONE COLONNE

                        //lstColumns.Append(new Column() { Min = 1, Max = 1, Width = 25, CustomWidth = true });
                        //lstColumns.Append(new Column() { Min = 2, Max = 2, Width = 9, CustomWidth = true });
                        //lstColumns.Append(new Column() { Min = 3, Max = 3, Width = 9, CustomWidth = true });

                        // Create custom widths for columns

                        if (_strCols != "")
                        {
                            Columns columns = worksheetPart.Worksheet.GetFirstChild<Columns>();
                            //Boolean needToInsertColumns = false;
                            if (columns == null)
                            {
                                columns = new Columns();
                                //needToInsertColumns = true;
                            }

                            //Columns columns = new Columns(
                            //        new Column // Id column
                            //        {
                            //            Min = 1,
                            //            Max = 1,
                            //            Width = 20,
                            //            CustomWidth = true
                            //        },
                            //        new Column // Name and Birthday columns
                            //        {
                            //            Min = 2,
                            //            Max = 2,
                            //            Width = 20,
                            //            CustomWidth = true
                            //        },
                            //        new Column // Salary column
                            //        {
                            //            Min = 3,
                            //            Max = 3,
                            //            Width = 10,
                            //            CustomWidth = true
                            //        });

                            //if (_intCols != null && _intCols.Length > 0)
                            //{
                            //    for (int iR = 0; i < _intCols.GetLength(0); i++)
                            //    {
                            //        UInt32 iMin = Convert.ToUInt32(_intCols[i, 0]);
                            //        UInt32 iMax = Convert.ToUInt32(_intCols[i, 1]);
                            //        int iWth = _intCols[i,2];

                            //        columns.Append(new Column() { Min = iMin, Max = iMax, Width = iWth, CustomWidth = true });
                            //    }
                            //}

                            Console.WriteLine("xxxx");

                            string[] c = _strCols.Split(';');
                            foreach (string ss in c)
                            {
                                if (ss.Trim() != "")
                                {
                                    Console.WriteLine("zzz");

                                    string[] cc = ss.Split(',');

                                    UInt32 iMin = Convert.ToUInt32(cc[0]);
                                    UInt32 iMax = Convert.ToUInt32(cc[1]);
                                    int iWth = Convert.ToInt32(cc[2]);

                                    Console.WriteLine("zzz");

                                    columns.Append(new Column() { Min = iMin, Max = iMax, Width = iWth, CustomWidth = true });
                                }
                            }

                            worksheetPart.Worksheet.AppendChild(columns);
                        }

                        Sheets sheets = workbookPart.Workbook.AppendChild(new Sheets());

                        Sheet sheet = new Sheet() { Id = workbookPart.GetIdOfPart(worksheetPart), SheetId = 1, Name = "Foglio1" };

                        sheets.Append(sheet);

                        workbookPart.Workbook.Save();

                        //List<Employee> employees = Employees.EmployeesList;

                        SheetData sheetData = worksheetPart.Worksheet.AppendChild(new SheetData());

                        var drawingsPart = worksheetPart.AddNewPart<DrawingsPart>();

                        //s = _clsDef.PATHPRNXLS + "Modelli\\Logo_Frassiflex2.jpg";
                        //XlsImmagine(worksheetPart, drawingsPart, s, 1, 1);

                        // Constructing header

                        Row row = new Row();
                        /* Righe vuote                    //Righe vuote di testata
                        row.Append();
                        sheetData.AppendChild(row);

                        for (int n = 0; n < 0; n++)
                        {
                            row = new Row();
                            row.Append();
                            sheetData.AppendChild(row);
                        }

                        row = new Row();
                        row.Append(ConstructCell(sDesArt, CellValues.String, 0));
                        sheetData.AppendChild(row);

                        row = new Row();
                        row.Append(ConstructCell("", CellValues.String, 0));
                        sheetData.AppendChild(row);
                        */

                        /*
                        row = new Row();
                        row.Append(
                            ConstructCell("", CellValues.String, 2),
                            ConstructCell("Id", CellValues.String, 2),
                            ConstructCell("Name", CellValues.String, 2),
                            ConstructCell("Birth Date", CellValues.String, 2),
                            ConstructCell("Salary", CellValues.String, 2));
                        // Insert the header row to the Sheet Data
                        sheetData.AppendChild(row);
                        */

                        // Inserting each employee
                        //foreach (var employee in employees)

                        Console.WriteLine("1111");


                        //row = new Row();
                        //row.Append(
                        //    ConstructCell("", CellValues.String, 2),
                        //    ConstructCell("Id", CellValues.String, 2),
                        //    ConstructCell("Name", CellValues.String, 2),
                        //    ConstructCell("Birth Date", CellValues.String, 2),
                        //    ConstructCell("Salary", CellValues.String, 2));

                        //// Insert the header row to the Sheet Data
                        //sheetData.AppendChild(row);

                        row = new Row();
                        s = (string)tabRow.Rows[ii]["art_des"];
                        row.Append(ConstructCell(s, CellValues.String, 2));

                        s = "PLU:" + (string)tabRow.Rows[ii]["art_plu"];
                        row.Append(ConstructCell(s, CellValues.String, 2));

                        //i = 0;
                        //foreach (DataColumn c in tabRow.Columns)
                        //{
                        //    i++;
                        //    if (i == 2)
                        //        row.Append(ConstructCell("", CellValues.String, 2));
                        //    else if (i > 2)
                        //        row.Append(ConstructCell(c.ColumnName, CellValues.String, 2));
                        //}

                        sheetData.AppendChild(row);


                        string sIng = "";
                        s = _clsDef.PATHINGREDIENTI + "et01_" + (string)tabRow.Rows[ii]["art_cod"] + ".rtf";
                        if (File.Exists(s))
                        {
                            RichTextBox rch = new RichTextBox();
                            rch.LoadFile(s);
                            sIng = rch.Text;
                            //sIng = _clsFun.StrSplit(sIng, 40);
                        }

                        if (sIng != "")
                        {
                            row = new Row();
                            row.Height = 50;
                            row.CustomHeight = true;
                            row.Append(ConstructCell(sIng, CellValues.String, 4));
                            row.Append(ConstructCell("", CellValues.String, 1));
                            //row.Append(ConstructCell("", CellValues.String, 4));
                            sheetData.AppendChild(row);
                        }

                        /*
                        foreach (DataRow y in tabRow.Rows)
                        {
                            row = new Row();
                            i = 0;
                            foreach (DataColumn c in tabRow.Columns)
                            {
                                i++;
                                //if (i == 2)
                                //    row.Append(ConstructCell("", CellValues.String, 1));
                                if (i >= 2)
                                {
                                    if (c.DataType.ToString() == "System.String")
                                    {
                                        //Cell cell = ConstructCell((string)y[c.ColumnName], CellValues.Number, 1);
                                        row.Append(ConstructCell((string)y[c.ColumnName], CellValues.String, 1));
                                    }
                                    else if (c.DataType.ToString() == "System.Decimal")
                                    {
                                        row.Append(ConstructCell(Convert.ToString(y[c.ColumnName]).Replace(",", "."), CellValues.Number, 1));
                                    }
                                }
                            }
                            //    row.Append(
                            //{
                            //        ConstructCell((string)y[c.ColumnName], CellValues.Number, 1)
                            //        //ConstructCell(employee.Name, CellValues.String, 1),
                            //        //ConstructCell(employee.DOB.ToString("yyyy/MM/dd"), CellValues.String, 1),
                            //        //ConstructCell(employee.Salary.ToString(), CellValues.Number, 1));
                            //}
                            sheetData.AppendChild(row);
                        }
                         * */
                        /**/
                        //if (File.Exists(sBloImg))
                        //    XlsImmagine(worksheetPart, drawingsPart, sBloImg, 30, 2);

                        worksheetPart.Worksheet.Save();
                    }

                    //string sFil = "C:\\APproject\\XLS\\" + DateTime.Now.ToString("yyyyMMddHHmm") + "_Listino_" + "_" + i.ToString() + ".xls";

                    Process.Start("Excel.exe", sFil);
                }
            }
        }

        public void xlsLibroIngredienti(DataTable tabRow, string strFil, Boolean bolPrv)
        {
            //_clsFun.ErrorLog("101", "passo");


            string sMsg = "";
            string _strCols = "1, 1, 80;";

            Boolean b = true;

            if (File.Exists(strFil))
            {
                try
                {
                    File.Delete(strFil);
                }
                catch (Exception ex)
                {
                    MessageBox.Show("File excel aperto!", "CONTROLLO FILE", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    b = false;
                }
            }

            if (b)
            {
                string sFil = strFil;

                _clsFun.ErrorLog("102", sFil);

                using (SpreadsheetDocument document = SpreadsheetDocument.Create(sFil, SpreadsheetDocumentType.Workbook))
                {


                    WorkbookPart workbookPart = document.AddWorkbookPart();
                    workbookPart.Workbook = new Workbook();

                    WorksheetPart worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
                    worksheetPart.Worksheet = new Worksheet();

                    // Adding style
                    WorkbookStylesPart stylePart = workbookPart.AddNewPart<WorkbookStylesPart>();
                    stylePart.Stylesheet = GenerateStylesheet();
                    stylePart.Stylesheet.Save();

                    // Setting up columns IMPOSTAZIONE COLONNE

                    //lstColumns.Append(new Column() { Min = 1, Max = 1, Width = 25, CustomWidth = true });
                    //lstColumns.Append(new Column() { Min = 2, Max = 2, Width = 9, CustomWidth = true });
                    //lstColumns.Append(new Column() { Min = 3, Max = 3, Width = 9, CustomWidth = true });

                    // Create custom widths for columns

                    if (_strCols != "")
                    {
                        Columns columns = worksheetPart.Worksheet.GetFirstChild<Columns>();
                        //Boolean needToInsertColumns = false;
                        if (columns == null)
                        {
                            columns = new Columns();
                            //needToInsertColumns = true;
                        }

                        //Columns columns = new Columns(
                        //        new Column // Id column
                        //        {
                        //            Min = 1,
                        //            Max = 1,
                        //            Width = 20,
                        //            CustomWidth = true
                        //        },
                        //        new Column // Name and Birthday columns
                        //        {
                        //            Min = 2,
                        //            Max = 2,
                        //            Width = 20,
                        //            CustomWidth = true
                        //        },
                        //        new Column // Salary column
                        //        {
                        //            Min = 3,
                        //            Max = 3,
                        //            Width = 10,
                        //            CustomWidth = true
                        //        });

                        //if (_intCols != null && _intCols.Length > 0)
                        //{
                        //    for (int iR = 0; i < _intCols.GetLength(0); i++)
                        //    {
                        //        UInt32 iMin = Convert.ToUInt32(_intCols[i, 0]);
                        //        UInt32 iMax = Convert.ToUInt32(_intCols[i, 1]);
                        //        int iWth = _intCols[i,2];

                        //        columns.Append(new Column() { Min = iMin, Max = iMax, Width = iWth, CustomWidth = true });
                        //    }
                        //}

                        Console.WriteLine("xxxx");

                        string[] c = _strCols.Split(';');
                        foreach (string ss in c)
                        {
                            if (ss.Trim() != "")
                            {
                                Console.WriteLine("zzz");

                                string[] cc = ss.Split(',');

                                UInt32 iMin = Convert.ToUInt32(cc[0]);
                                UInt32 iMax = Convert.ToUInt32(cc[1]);
                                int iWth = Convert.ToInt32(cc[2]);

                                Console.WriteLine("zzz");

                                columns.Append(new Column() { Min = iMin, Max = iMax, Width = iWth, CustomWidth = true });
                            }
                        }

                        worksheetPart.Worksheet.AppendChild(columns);
                    }




                    Sheets sheets = workbookPart.Workbook.AppendChild(new Sheets());

                    Sheet sheet = new Sheet() { Id = workbookPart.GetIdOfPart(worksheetPart), SheetId = 1, Name = "Foglio1" };

                    sheets.Append(sheet);

                    workbookPart.Workbook.Save();

                    //List<Employee> employees = Employees.EmployeesList;

                    SheetData sheetData = worksheetPart.Worksheet.AppendChild(new SheetData());

                    var drawingsPart = worksheetPart.AddNewPart<DrawingsPart>();

                    //s = _clsDef.PATHPRNXLS + "Modelli\\Logo_Frassiflex2.jpg";
                    //XlsImmagine(worksheetPart, drawingsPart, s, 1, 1);

                    // Constructing header


                    for (int ii = 0; ii <= tabRow.Rows.Count - 1; ii++)
                    {
                        string s = "";

                        string sDesArt = "";

                        int i = 0;





                        Row row = new Row();
                        /* Righe vuote                    //Righe vuote di testata
                        row.Append();
                        sheetData.AppendChild(row);

                        for (int n = 0; n < 0; n++)
                        {
                            row = new Row();
                            row.Append();
                            sheetData.AppendChild(row);
                        }

                        row = new Row();
                        row.Append(ConstructCell(sDesArt, CellValues.String, 0));
                        sheetData.AppendChild(row);

                        row = new Row();
                        row.Append(ConstructCell("", CellValues.String, 0));
                        sheetData.AppendChild(row);
                        */

                        /*
                        row = new Row();
                        row.Append(
                            ConstructCell("", CellValues.String, 2),
                            ConstructCell("Id", CellValues.String, 2),
                            ConstructCell("Name", CellValues.String, 2),
                            ConstructCell("Birth Date", CellValues.String, 2),
                            ConstructCell("Salary", CellValues.String, 2));
                        // Insert the header row to the Sheet Data
                        sheetData.AppendChild(row);
                        */

                        // Inserting each employee
                        //foreach (var employee in employees)

                        Console.WriteLine("1111");


                        //row = new Row();
                        //row.Append(
                        //    ConstructCell("", CellValues.String, 2),
                        //    ConstructCell("Id", CellValues.String, 2),
                        //    ConstructCell("Name", CellValues.String, 2),
                        //    ConstructCell("Birth Date", CellValues.String, 2),
                        //    ConstructCell("Salary", CellValues.String, 2));

                        //// Insert the header row to the Sheet Data
                        //sheetData.AppendChild(row);

                        row = new Row();
                        s = (string)tabRow.Rows[ii]["art_des"];

                        if (bolPrv)
                        {
                            s = (s + new string(' ', 50)).Substring(0,50);
                            s += "Prezzo: " + ((decimal)tabRow.Rows[ii]["art_prv"]).ToString("#,##0.00");
                        }

                        row.Append(ConstructCell(s, CellValues.String, 2));

                        s = "PLU:" + (string)tabRow.Rows[ii]["art_plu"];
                        row.Append(ConstructCell(s, CellValues.String, 2));

                        //i = 0;
                        //foreach (DataColumn c in tabRow.Columns)
                        //{
                        //    i++;
                        //    if (i == 2)
                        //        row.Append(ConstructCell("", CellValues.String, 2));
                        //    else if (i > 2)
                        //        row.Append(ConstructCell(c.ColumnName, CellValues.String, 2));
                        //}

                        sheetData.AppendChild(row);


                        string sIng = "";
                        s = _clsDef.PATHINGREDIENTI + "et01_" + (string)tabRow.Rows[ii]["art_cod"] + ".rtf";
                        if (File.Exists(s))
                        {
                            RichTextBox rch = new RichTextBox();
                            rch.LoadFile(s);
                            sIng = rch.Text;
                            //sIng = _clsFun.StrSplit(sIng, 40);
                        }

                        if (sIng != "")
                        {
                            row = new Row();
                            row.Height = 50;
                            row.CustomHeight = true;
                            row.Append(ConstructCell(sIng, CellValues.String, 4));
                            row.Append(ConstructCell("", CellValues.String, 1));
                            //row.Append(ConstructCell("", CellValues.String, 4));
                            sheetData.AppendChild(row);
                        }

                        /*
                        foreach (DataRow y in tabRow.Rows)
                        {
                            row = new Row();
                            i = 0;
                            foreach (DataColumn c in tabRow.Columns)
                            {
                                i++;
                                //if (i == 2)
                                //    row.Append(ConstructCell("", CellValues.String, 1));
                                if (i >= 2)
                                {
                                    if (c.DataType.ToString() == "System.String")
                                    {
                                        //Cell cell = ConstructCell((string)y[c.ColumnName], CellValues.Number, 1);
                                        row.Append(ConstructCell((string)y[c.ColumnName], CellValues.String, 1));
                                    }
                                    else if (c.DataType.ToString() == "System.Decimal")
                                    {
                                        row.Append(ConstructCell(Convert.ToString(y[c.ColumnName]).Replace(",", "."), CellValues.Number, 1));
                                    }
                                }
                            }
                            //    row.Append(
                            //{
                            //        ConstructCell((string)y[c.ColumnName], CellValues.Number, 1)
                            //        //ConstructCell(employee.Name, CellValues.String, 1),
                            //        //ConstructCell(employee.DOB.ToString("yyyy/MM/dd"), CellValues.String, 1),
                            //        //ConstructCell(employee.Salary.ToString(), CellValues.Number, 1));
                            //}
                            sheetData.AppendChild(row);
                        }
                            * */
                        /**/
                        //if (File.Exists(sBloImg))
                        //    XlsImmagine(worksheetPart, drawingsPart, sBloImg, 30, 2);

                        worksheetPart.Worksheet.Save();
                    }

                    //string sFil = "C:\\APproject\\XLS\\" + DateTime.Now.ToString("yyyyMMddHHmm") + "_Listino_" + "_" + i.ToString() + ".xls";

                    try
                    {
                        Process.Start("Excel.exe", sFil);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Excel non attivo!", "CONTROLLO FILE", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    }

                }
            }

        }

        private Cell ConstructCell(string value, CellValues dataType, uint styleIndex = 0)
        {
            return new Cell()
            {
                CellValue = new CellValue(value),
                DataType = new EnumValue<CellValues>(dataType),

                StyleIndex = styleIndex
            };
        }

        private Stylesheet GenerateStylesheet()
        {
            Stylesheet styleSheet = null;


            Fonts fonts = new Fonts(
                    new Font(                       // Index 0 - default
                    new FontSize() { Val = 10 }
                    ),
                    new Font(                       // Index descrizione e plu
                    new FontSize() { Val = 11 },
                    new Bold(),
                    new Color() { Rgb = "000000" }
                    ),
                    new Font(                       // Index 2 - header
                    new FontSize() { Val = 10 },
                    new Bold(),
                    new Color() { Rgb = "000000" }
                    )
                );

            Fills fills = new Fills(
                    new Fill(new PatternFill() { PatternType = PatternValues.None }), // Index 0 - default
                    new Fill(new PatternFill() { PatternType = PatternValues.Gray125 }), // Index 1 - default
                    new Fill(new PatternFill(new ForegroundColor { Rgb = new HexBinaryValue() { Value = "FFFF00" } }) { PatternType = PatternValues.Solid }) // Index 2 - header
                );

            Borders borders = new Borders(
                    new Border(), // index 0 default
                    new Border( // index 1 black border
                        new LeftBorder(new Color() { Auto = true }) { Style = BorderStyleValues.None },
                        new RightBorder(new Color() { Auto = true }) { Style = BorderStyleValues.Medium },
                        new TopBorder(new Color() { Auto = true }) { Style = BorderStyleValues.None },
                        new BottomBorder(new Color() { Auto = true }) { Style = BorderStyleValues.Medium },
                        new DiagonalBorder()),
                    new Border( // index 1 black border
                        new LeftBorder(new Color() { Auto = true }) { Style = BorderStyleValues.None },
                        new RightBorder(new Color() { Auto = true }) { Style = BorderStyleValues.Medium },
                        new TopBorder(new Color() { Auto = true }) { Style = BorderStyleValues.None },
                        new BottomBorder(new Color() { Auto = true }) { Style = BorderStyleValues.Medium },
                        new DiagonalBorder()),
                    new Border( // index 1 black border
                        new LeftBorder(new Color() { Auto = true }) { Style = BorderStyleValues.None },
                        new RightBorder(new Color() { Auto = true }) { Style = BorderStyleValues.None },
                        new TopBorder(new Color() { Auto = true }) { Style = BorderStyleValues.None },
                        new BottomBorder(new Color() { Auto = true }) { Style = BorderStyleValues.Medium },
                        new DiagonalBorder())
                );


              //Alignment alignment1 = new Alignment(){ WrapText = true };
              //cellFormat2.Append(alignment1);


            CellFormats cellFormats = new CellFormats(
                    new CellFormat(), // default
                    new CellFormat { FontId = 0, FillId = 0, BorderId = 2, ApplyBorder = true }, // body
                    new CellFormat { FontId = 1, FillId = 2, BorderId = 1, ApplyFill = true }, // header
                    new CellFormat { FontId = 0, FillId = 0, BorderId = 2, ApplyBorder = true }
                );


            CellFormat cellFormat2 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)0U, FillId = (UInt32Value)0U, BorderId = 3, FormatId = (UInt32Value)0U, ApplyAlignment = true };
            Alignment alignment1 = new Alignment() { WrapText = true };
            cellFormat2.Append(alignment1);

            cellFormats.Append(cellFormat2);

            styleSheet = new Stylesheet(fonts, fills, borders, cellFormats);

            return styleSheet;
        }

        private void XlsImmagine(WorksheetPart worksheetPart, DrawingsPart drawingsPart, string imageFileName, int rowNumber, int colNumber)
        {
            //var drawingsPart = worksheetPart.AddNewPart<DrawingsPart>();

            if (!worksheetPart.Worksheet.ChildElements.OfType<Drawing>().Any())
            {
                worksheetPart.Worksheet.Append(new Drawing { Id = worksheetPart.GetIdOfPart(drawingsPart) });
            }

            if (drawingsPart.WorksheetDrawing == null)
            {
                drawingsPart.WorksheetDrawing = new WorksheetDrawing();
            }

            var worksheetDrawing = drawingsPart.WorksheetDrawing;

            var imagePart = drawingsPart.AddImagePart(ImagePartType.Jpeg);

            using (var stream = new FileStream(imageFileName, FileMode.Open))
            {
                imagePart.FeedData(stream);
            }

            System.Drawing.Bitmap bm = new System.Drawing.Bitmap(imageFileName);
            DocumentFormat.OpenXml.Drawing.Extents extents = new DocumentFormat.OpenXml.Drawing.Extents();
            var extentsCx = (long)bm.Width * (long)((float)914400 / bm.HorizontalResolution);
            var extentsCy = (long)bm.Height * (long)((float)914400 / bm.VerticalResolution);
            bm.Dispose();

            var colOffset = 0;
            var rowOffset = 0;
            //int colNumber = 1;
            //int rowNumber = 1;

            var nvps = worksheetDrawing.Descendants<Xdr.NonVisualDrawingProperties>();
            var nvpId = nvps.Count() > 0 ?
                (UInt32Value)worksheetDrawing.Descendants<Xdr.NonVisualDrawingProperties>().Max(p => p.Id.Value) + 1 :
                1U;

            var oneCellAnchor = new Xdr.OneCellAnchor(
                new Xdr.FromMarker
                {
                    ColumnId = new Xdr.ColumnId((colNumber - 1).ToString()),
                    RowId = new Xdr.RowId((rowNumber - 1).ToString()),
                    ColumnOffset = new Xdr.ColumnOffset(colOffset.ToString()),
                    RowOffset = new Xdr.RowOffset(rowOffset.ToString())
                },
                new Xdr.Extent { Cx = extentsCx, Cy = extentsCy },
                new Xdr.Picture(
                    new Xdr.NonVisualPictureProperties(
                        new Xdr.NonVisualDrawingProperties { Id = nvpId, Name = "Picture " + nvpId, Description = imageFileName },
                        new Xdr.NonVisualPictureDrawingProperties(new A.PictureLocks { NoChangeAspect = true })
                    ),
                    new Xdr.BlipFill(
                        new A.Blip { Embed = drawingsPart.GetIdOfPart(imagePart), CompressionState = A.BlipCompressionValues.Print },
                        new A.Stretch(new A.FillRectangle())
                    ),
                    new Xdr.ShapeProperties(
                        new A.Transform2D(
                            new A.Offset { X = 0, Y = 0 },
                            new A.Extents { Cx = extentsCx, Cy = extentsCy }
                        ),
                        new A.PresetGeometry { Preset = A.ShapeTypeValues.Rectangle }
                    )
                ),
                new Xdr.ClientData()
            );

            worksheetDrawing.Append(oneCellAnchor);
        }


    }

}
