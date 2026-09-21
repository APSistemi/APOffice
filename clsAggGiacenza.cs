using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Data;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace APOffice
{
    class clsAggGiacenza
    {
        clsDefine _clsDef = new clsDefine();
        clsFuncs _clsFun = new clsFuncs();

        private string _strConSql = "";
        private string _strConSqlSta = "";

        private DataTable _tabGia = new DataTable();
        private DataTable _tabInv = new DataTable();
        private DateTime _dayInv = new DateTime(2015, 1, 1);

        public ProgressBar _proBar = new ProgressBar();
        public Label _lblMsg = new Label();

        public clsAggGiacenza()
        {
            _strConSql = _clsFun.ConSql("");
            _strConSqlSta = _clsFun.ConSql("3");

            string s = "DELETE FROM AnaArtGiacenza";
            _clsFun.SqlWrite(s, _strConSql);

            s = "SELECT * FROM AnaArtGiacenza";
            _tabGia = _clsFun.FillTabSql("AnaArtGiacenza", s, true, _strConSql);
            DataColumn[] keys = new DataColumn[1];
            keys[0] = _tabGia.Columns["gia_art"];
            _tabGia.PrimaryKey = keys;
        }

        public void AggGiacenza()
        {
            DataInvTotale();
            _lblMsg.Text = "Lettura inventario";
            FillInventario();
            _lblMsg.Text = "Lettura fatture";
            FillMovimenti("FAT");
            _lblMsg.Text = "Lettura movimenti";
            FillMovimenti("MOV");
            _lblMsg.Text = "Lettura venduto";
            FillVendite();
            _lblMsg.Text = "Aggiornamento giacenza";
            GiaAggiorna();
        }

        private void DataInvTotale()
        {
            DataRow x;
            DataRow[] j;

            string s = "SELECT * FROM GesInvTestate WHERE int_tip='T' ORDER BY int_day DESC";
            DataTable t = _clsFun.FillTabSql("GesInvTestate", s, true, _strConSql);

            if(t.Rows.Count > 0)
                _dayInv = (DateTime)t.Rows[0]["int_day"];
        }

        private void FillInventario()
        {
            //string s = "SELECT art_cod, art_umi FROM AnaArticoli WHERE art_sta='" + _clsDef.STAATT + "'";
            string s = "SELECT art_cod, art_umi FROM AnaArticoli";
            DataTable tArt = _clsFun.FillTabSql("AnaArticoli", s, false, _strConSql);
            DataTable tInv = new DataTable("TabInv");
            DataRow x;
            DataRow[] j;

            _proBar.Value = 0;
            _proBar.Maximum = tArt.Rows.Count;
            _proBar.Minimum = 0;

            foreach(DataRow y in tArt.Rows)
            {
                _proBar.Increment(1);
                Application.DoEvents();

                if ((string)y["art_cod"] == "0023728")
                    Console.WriteLine("xxxx");

                s = "SELECT inv_art, inv_qta, inv_cos, int_day, int_num FROM GesInventario ";
                s += "LEFT JOIN GesInvTestate ON GesInventario.inv_num = GesInvTestate.int_num ";
                s += "WHERE int_day >= " + _clsFun.DaySql(_dayInv) + " AND int_tip<>'G' AND inv_art='" + y["art_cod"] + "' ";
                s += "ORDER BY int_day DESC";
                tInv = _clsFun.FillTabSql("GesInventatio", s, false, _strConSql);

                if(tInv.Rows.Count > 0)
                {
                    string sInvNum = (string)tInv.Rows[0]["int_num"];

                    j = _tabGia.Select("gia_art='" + y["art_cod"] + "'");
                    if(j.Length == 0)
                    {
                        _tabGia.Rows.Add(InsArticolo(_tabGia, y, (DateTime)tInv.Rows[0]["int_day"]));
                        j = _tabGia.Select("gia_art='" + y["art_cod"] + "'");
                    }

                    foreach (DataRow yy in tInv.Rows)
                    {
                        if ((string)yy["int_num"] == sInvNum)
                        {
                            //j[0]["gia_iqt"] = (decimal)j[0]["gia_iqt"] + (decimal)tInv.Rows[0]["inv_qta"];
                            j[0]["gia_iqt"] = (decimal)j[0]["gia_iqt"] + (decimal)yy["inv_qta"];
                            j[0]["gia_ico"] = (decimal)yy["inv_cos"];
                            j[0]["gia_dti"] = (DateTime)yy["int_day"];
                        }
                    }
                }
                else
                {
                    j = _tabGia.Select("gia_art='" + y["art_cod"] + "'");
                    if (j.Length > 0)
                    {
                        j[0]["gia_iqt"] = 0;
                        j[0]["gia_ico"] = 0;
                        j[0]["gia_dti"] = _dayInv;
                    }
                }
             }
        }

        private DataRow InsArticolo(DataTable tabArt, DataRow rowArt, DateTime dayInv)
        {
            DataRow x = tabArt.NewRow();
            x["gia_neg"] = "";
            x["gia_mag"] = "";
            x["gia_ora"] = "";
            x["gia_day"] = dayInv;
            x["gia_art"] = rowArt["art_cod"];
            x["gia_umi"] = rowArt["art_umi"];
            x["gia_ard"] = "";

            x["gia_iqt"] = 0;
            x["gia_ico"] = 0;
            x["gia_mqt"] = 0;
            x["gia_aqt"] = 0;
            x["gia_aco"] = 0;
            x["gia_vqt"] = 0;
            x["gia_vpv"] = 0;

            x["gia_dti"] = _dayInv;             //Data inventario
            return x;
        }

        private void FillMovimenti(string strTip)
        {
            DateTime dDay = _dayInv;
            string s = "";
            DataRow[] j;
            int iSgn = 1;

            s = "SELECT art_cod, art_umi FROM AnaArticoli WHERE art_sta='" + _clsDef.STAATT + "'";
            DataTable tArt = _clsFun.FillTabSql("AnaArticoli", s, false, _strConSql);
            DataColumn[] keys = new DataColumn[1];
            keys[0] = tArt.Columns["art_cod"];
            tArt.PrimaryKey = keys;

            if (strTip == "FAT")
            {
                s = "SELECT ";
                s += "GesFatTestate.fat_tpd, ";
                s += "GesFatTestate.fat_ann, ";
                s += "GesFatTestate.fat_cfo, ";
                s += "GesMovimenti.mov_rfa, ";
                s += "GesFatTestate.fat_ddo, ";
                s += "GesMovimenti.mov_art, ";
                s += "GesMovimenti.mov_ard, ";
                s += "GesMovimenti.mov_cos, ";
                s += "GesMovimenti.mov_prv, ";
                s += "GesMovimenti.mov_imp, ";
                s += "GesMovimenti.mov_qta, ";
                s += "GesMovimenti.mov_qkg ";
                s += "FROM GesFatTestate ";
                s += "LEFT OUTER JOIN GesMovimenti ON GesFatTestate.fat_yfa = GesMovimenti.mov_yfa AND GesFatTestate.fat_nfa = GesMovimenti.mov_nfa ";
                s += "WHERE ";
                s += "GesFatTestate.fat_ddo >= " + _clsFun.DaySql(dDay) + " AND ";
                s += "GesMovimenti.mov_ori = '' ";
                s += "ORDER BY GesMovimenti.mov_art";
            }
            else
            {
                s = "SELECT ";
                s += "GesMovTestate.mot_neg, ";
                s += "GesMovTestate.mot_yfa, ";
                s += "GesMovTestate.mot_nfa, ";
                s += "GesMovTestate.mot_ddo, ";
                s += "GesMovTestate.mot_ndo, ";
                s += "GesMovTestate.mot_cfo, ";
                s += "GesMovimenti.mov_ymo, ";
                s += "GesMovimenti.mov_nmo, ";
                s += "GesMovimenti.mov_art, ";
                s += "GesMovimenti.mov_umi, ";
                s += "GesMovimenti.mov_qta, ";
                s += "GesMovimenti.mov_qkg, ";
                s += "GesMovimenti.mov_cos, ";
                s += "GesMovimenti.mov_prv, ";
                s += "GesMovimenti.mov_imp, ";
                s += "GesMovimenti.mov_ori, ";
                s += "TabMovCausali.tab_sgm, ";
                s += "TabMovCausali.tab_des, ";
                s += "TabMovCausali.tab_cfo, ";
                s += "TabMovCausali.tab_tip ";
                s += "FROM GesMovTestate LEFT OUTER JOIN ";
                s += "GesMovimenti ON GesMovTestate.mot_ymo = GesMovimenti.mov_ymo AND GesMovTestate.mot_nmo = GesMovimenti.mov_nmo LEFT OUTER JOIN ";
                s += "TabMovCausali ON GesMovTestate.mot_cau = TabMovCausali.tab_cod ";
                s += "WHERE "; //mov_art='" + _strArtCod + "' AND ";
                s += "GesMovTestate.mot_ddo >= " + _clsFun.DaySql(dDay) + " ";
            }
            DataTable tMov = _clsFun.FillTabSql("GesMovimenti", s, false, _strConSql);

            _proBar.Value = 0;
            _proBar.Maximum = tMov.Rows.Count;
            _proBar.Minimum = 0;

            foreach (DataRow y in tMov.Rows)
            {
                _proBar.Increment(1);
                Application.DoEvents();

                if ((string)y["mov_art"] == "0010435")
                    Console.WriteLine("xxxx");

                j = _tabGia.Select("gia_art='" + (string)y["mov_art"] + "'");
                if (j.Length == 0)
                {
                    j = tArt.Select("art_cod='" + (string)y["mov_art"] + "'");
                    if(j.Length > 0)
                    {
                        _tabGia.Rows.Add(InsArticolo(_tabGia, j[0], dDay));
                        j = _tabGia.Select("gia_art='" + y["mov_art"] + "'");
                    }
                }

                if (j.Length > 0)
                {
                    DateTime dd = DateTime.Now;
                    if (strTip == "FAT")
                        dd = (DateTime)y["fat_ddo"];
                    else
                        dd = (DateTime)y["mot_ddo"];

                    TimeSpan tm = ((DateTime)j[0]["gia_dti"]).Subtract(dd);

                    if ( tm.Days <= 0)
                    {
                        iSgn = 1;
                        if (strTip == "FAT")
                        {
                            if ((string)tMov.Rows[0]["fat_tpd"] == "FV")
                                iSgn = 1;
                        }
                        else
                        {
                            if ((string)y["tab_sgm"] == "-")
                                iSgn = -1;
                        }

                        if (j.Length > 0 && iSgn != 0)
                        {
                            if (strTip == "FAT")
                            {
                                if ((string)j[0]["gia_umi"] == "KG")
                                    j[0]["gia_aqt"] = (decimal)j[0]["gia_aqt"] + ((decimal)y["mov_qkg"] * iSgn);
                                else
                                    j[0]["gia_aqt"] = (decimal)j[0]["gia_aqt"] + ((decimal)y["mov_qta"] * iSgn);

                                j[0]["gia_aco"] = (decimal)y["mov_cos"];
                            }
                            else
                            {
                                if (((string)y["tab_tip"]).Contains("D"))    //Movimento tipo documento)
                                {
                                    if((string)y["tab_cfo"] == "FOR")
                                    {
                                        if ((string)j[0]["gia_umi"] == "KG")
                                            j[0]["gia_aqt"] = (decimal)j[0]["gia_aqt"] + ((decimal)y["mov_qkg"] * iSgn);
                                        else
                                            j[0]["gia_aqt"] = (decimal)j[0]["gia_aqt"] + ((decimal)y["mov_qta"] * iSgn);

                                        j[0]["gia_aco"] = (decimal)y["mov_cos"];
                                    }
                                    else if ((string)y["tab_cfo"] == "CLI")
                                    {
                                        if ((string)j[0]["gia_umi"] == "KG")
                                            j[0]["gia_vqt"] = (decimal)j[0]["gia_vqt"] + ((decimal)y["mov_qkg"] * iSgn);
                                        else
                                            j[0]["gia_vqt"] = (decimal)j[0]["gia_vqt"] + ((decimal)y["mov_qta"] * iSgn);
                                    }
                                }
                                else
                                {

                                    if ((string)j[0]["gia_umi"] == "KG")
                                        j[0]["gia_mqt"] = (decimal)j[0]["gia_mqt"] + ((decimal)y["mov_qkg"] * iSgn);
                                    else
                                        j[0]["gia_mqt"] = (decimal)j[0]["gia_mqt"] + ((decimal)y["mov_qta"] * iSgn);
                                }
                            }
                        }
                    }
                }
            }
        }

        private void FillVendite()
        {
            string s = "SELECT art_cod, art_umi FROM AnaArticoli WHERE art_sta='" + _clsDef.STAATT + "'";
            DataTable tArt = _clsFun.FillTabSql("AnaArticoli", s, false, _strConSql);
            DataTable t = new DataTable("TabInv");
            DataRow x;
            DataRow[] j;
            int iSgn = 1;

            _proBar.Value = 0;
            _proBar.Maximum = tArt.Rows.Count;
            _proBar.Minimum = 0;

            foreach (DataRow y in tArt.Rows)
            {
                _proBar.Increment(1);
                Application.DoEvents();

                if ((string)y["art_cod"] == "0000035")
                    Console.WriteLine("xxxx");

                DateTime dDay = _dayInv;
                j = _tabGia.Select("gia_art='" + y["art_cod"] + "'");
                if(j.Length > 0)
                    dDay = (DateTime)j[0]["gia_dti"];

                s = "SELECT ";
                s += "ven_art, ";
                s += "SUM(ven_qta) AS ven_qta, ";
                s += "SUM(ven_qkg) AS ven_qkg, ";
                s += "MAX(ven_prz) AS ven_prz ";
                s += "FROM GesNegVen ";
                s += "WHERE (ven_day >= " + _clsFun.DaySql(dDay) + ") AND ven_art='" + y["art_cod"] + "' ";
                s += "GROUP BY ven_art";
                t = _clsFun.FillTabSql("GesVendite", s, true, _strConSqlSta);

                if (t.Rows.Count > 0)
                {
                    j = _tabGia.Select("gia_art='" + y["art_cod"] + "'");
                    if (j.Length == 0)
                    {
                        _tabGia.Rows.Add(InsArticolo(_tabGia, y, dDay));
                        j = _tabGia.Select("gia_art='" + y["art_cod"] + "'");
                    }

                    if (iSgn != 0)
                    {
                        if ((string)j[0]["gia_umi"] == "KG")
                            j[0]["gia_vqt"] = (decimal)j[0]["gia_vqt"] + ((decimal)t.Rows[0]["ven_qkg"] * iSgn);
                        else
                            j[0]["gia_vqt"] = (decimal)j[0]["gia_vqt"] + ((decimal)t.Rows[0]["ven_qta"] * iSgn);
                    }

                    j[0]["gia_vpv"] = (decimal)t.Rows[0]["ven_prz"];
                }
                else
                {
                    j = _tabGia.Select("gia_art='" + y["art_cod"] + "'");
                    if (j.Length > 0)
                    {
                        j[0]["gia_vqt"] = 0;
                        j[0]["gia_vpv"] = 0;
                    }
                }
            }
        }

        private void GiaAggiorna()
        {
            string s = "";
            DataRow x;
            DataRow[] j;
            decimal d = 0;

            //string sNeg = _clsFun.FileIni("R", clsDefine.enuIni.Ini006_codNegozio, "");
            string sNeg = _clsFun.FileIni("R", clsDefine.enuIni.Ini09CodiceAzienda, "");
            DateTime dDay = DateTime.Today;
            string sOra = DateTime.Now.ToString("HHmm");

            string sTabGia = "AnaArtGiacenza";

            s = "SELECT * FROM AnaArtGiacenza";
            DataTable tGia = _clsFun.FillTabSql(sTabGia, s, false, _strConSql);
            DataColumn[] keys = new DataColumn[3];
            keys[0] = tGia.Columns["gia_neg"];
            keys[1] = tGia.Columns["gia_mag"];
            keys[2] = tGia.Columns["gia_art"];
            tGia.PrimaryKey = keys;

            ArrayList aWhe = new ArrayList();
            aWhe.Add("gia_neg");
            aWhe.Add("gia_mag");
            aWhe.Add("gia_art");
            ArrayList aExl = new ArrayList();
            aExl.Add("gia_day");
            aExl.Add("gia_ora");

            _proBar.Value = 0;
            _proBar.Maximum = _tabGia.Rows.Count;
            _proBar.Minimum = 0;

            foreach (DataRow y in _tabGia.Rows)
            {
                _proBar.Increment(1);
                Application.DoEvents();

                if((string)y["gia_art"] == "0000100")
                    Console.WriteLine("zzzz");

                d = (decimal)y["gia_iqt"] + (decimal)y["gia_aqt"] + (decimal)y["gia_mqt"] - (decimal)y["gia_vqt"];
                //if((string)y["gia_umi"] == "KG")
                //    d = (decimal)y["gia_ikg"] + (decimal)y["gia_akg"] + (decimal)y["gia_mkg"] - (decimal)y["gia_vkg"];

                s = "UPDATE AnaArticoli SET art_gia=" + d.ToString().Replace(",", ".") + " WHERE art_cod='" + y["gia_art"] + "'";
                _clsFun.SqlWrite(s, _strConSql);

                x = tGia.NewRow();
                x["gia_neg"] = sNeg;
                x["gia_mag"] = "00";
                x["gia_day"] = dDay;
                x["gia_ora"] = sOra;
                x["gia_art"] = y["gia_art"];
                x["gia_ard"] = y["gia_ard"];
                x["gia_umi"] = y["gia_umi"];
                x["gia_iqt"] = y["gia_iqt"];
                x["gia_ico"] = y["gia_ico"];
                x["gia_mqt"] = y["gia_mqt"]; 
                x["gia_aqt"] = y["gia_aqt"];
                x["gia_aco"] = y["gia_aco"];
                x["gia_vqt"] = y["gia_vqt"];
                x["gia_vpv"] = y["gia_vpv"];
                x["gia_gia"] = d;

                j = tGia.Select("gia_neg='" + sNeg + "' AND gia_mag='00' AND gia_art='" + y["gia_art"] + "'");
                if(j.Length > 0)
                {
                    s = _clsFun.SqlUpdRow(sTabGia, tGia, j[0], x, aWhe, aExl);
                }
                else
                {
                    s = _clsFun.SqlInsertRow(sTabGia, tGia, x);
                }

                if (s != "")
                    _clsFun.SqlWrite(s, _strConSql);
            }
        }
    }
}
