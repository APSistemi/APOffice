using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.Data.SqlClient;
using System.Windows.Forms;
using System.Drawing;

namespace APOffice
{
    class clsGesGraph
    {
        clsFuncs _clsFun = new clsFuncs();

        private DataTable _tabFon = new DataTable();

        private string strGraph = "";

        public clsGesGraph()
        {
            string strConSql = _clsFun.ConSql("");
            string s = "SELECT * FROM TabFonts";
            _tabFon = _clsFun.FillTabSql("TabFonts", s, false, strConSql);
            strGraph = _clsFun.ParGet(clsDefine.enuParametri.ParGesGraph, strConSql);
        }
        
        public void SetGraph(Control ctrl, int intLiv)
        {
            if (strGraph == "S")
            {
                DataRow[] j;
                string s = "";

                s = ctrl.GetType().ToString();

                if (intLiv == 0)
                {
                    j = _tabFon.Select("tab_cod='001'");
                    try
                    {
                        if (j.Length > 0 && ((string)j[0]["tab_col"]).Trim() != "")
                        {
                            //c.BackColor = Color.FromName((string)j[0]["tab_col"]);
                            s = (string)j[0]["tab_col"];
                            ctrl.BackColor = System.Drawing.ColorTranslator.FromHtml("#" + s);
                        }
                    }
                    catch (Exception ex)
                    {
                        _clsFun.ErrorLog(ex.Message, Convert.ToString(j[0]["tab_cod"]));
                    }
                }

                intLiv++;

                foreach (Control c in ctrl.Controls)
                {
                    try
                    {
                        s = c.GetType().ToString();

                        if (c.GetType() == typeof(MenuStrip))
                        {
                            j = _tabFon.Select("tab_cod='002'");
                            if (j.Length > 0 && ((string)j[0]["tab_col"]).Trim() != "")
                            {
                                //c.BackColor = Color.FromName((string)j[0]["tab_col"]);
                                s = (string)j[0]["tab_col"];
                                c.BackColor = System.Drawing.ColorTranslator.FromHtml("#" + s);
                            }
                        }

                        else if (c.GetType() == typeof(Button))
                        {
                            j = _tabFon.Select("tab_cod='003'");
                            if (j.Length > 0 && ((string)j[0]["tab_col"]).Trim() != "")
                            {
                                //c.BackColor = Color.FromName((string)j[0]["tab_col"]);
                                s = (string)j[0]["tab_col"];
                                c.BackColor = System.Drawing.ColorTranslator.FromHtml("#" + s);
                            }
                        }

                        else if (c.GetType() == typeof(Panel))
                        {
                            j = _tabFon.Select("tab_cod='004'");
                            if (j.Length > 0 && ((string)j[0]["tab_col"]).Trim() != "")
                            {
                                //c.BackColor = Color.FromName((string)j[0]["tab_col"]);
                                s = (string)j[0]["tab_col"];
                                c.BackColor = System.Drawing.ColorTranslator.FromHtml("#" + s);

                                if (c.HasChildren)
                                {
                                    SetGraph(c, intLiv);
                                }
                            }
                        }

                        else if (c.GetType() == typeof(GroupBox))
                        {
                            j = _tabFon.Select("tab_cod='005'");
                            if (j.Length > 0 && ((string)j[0]["tab_col"]).Trim() != "")
                            {
                                //c.BackColor = Color.FromName((string)j[0]["tab_col"]);
                                s = (string)j[0]["tab_col"];
                                c.BackColor = System.Drawing.ColorTranslator.FromHtml("#" + s);

                                if (c.HasChildren)
                                {
                                    SetGraph(c, intLiv);
                                }
                            }
                        }

                        else if (c.GetType() == typeof(DataGridView))
                        {
                            j = _tabFon.Select("tab_cod='006'");
                            if (j.Length > 0 && ((string)j[0]["tab_col"]).Trim() != "")
                            {
                                //c.BackColor = Color.FromName((string)j[0]["tab_col"]);
                                s = (string)j[0]["tab_col"];

                                DataGridView d = (DataGridView)c;

                                d.BackgroundColor = System.Drawing.ColorTranslator.FromHtml("#" + s);

                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        _clsFun.ErrorLog(ex.Message, c.Name);
                    }
                }
            }
        }

        public Font Str2Font(Font fon, string sFont)
        {
            Font fo = fon;
            try
            {
                char[] chrArray = new char[] { ':' };
                string[] parts = sFont.Split(chrArray);
                if ((int)parts.Length < 3)
                {
                    throw new ArgumentException("Not a valid font string", "font");
                }
                fo = new Font(parts[0], float.Parse(parts[1]), (FontStyle)int.Parse(parts[2]));
            }
            catch (Exception exception)
            {
                Exception e = exception;
                Console.WriteLine("The file could not be read:");
                MessageBox.Show(e.Message);
            }
            Font font = fo;
            return font;
        }

    }
}
