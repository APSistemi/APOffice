using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.Data.OleDb;

namespace APOffice
{
    class clsGenMdb
    {
        clsDefine _clsDef = new clsDefine();
        clsFuncs _clsFun = new clsFuncs();

        private string _strConMdb = "";

        public clsGenMdb()
        {
            _strConMdb = _clsFun.ConMdb("C:\\APproject\\dBase\\DataPrn.mdb");
        }

        public void SalvaMov(DataTable tabTab)
        {
            string s = "DELETE FROM DocMov";
            _clsFun.MdbWrite(s, _strConMdb);

            DataTable t = _clsFun.FillTabMdb("DocMov", "SELECT * FROM DocMov", true, _strConMdb);

            foreach (DataRow y in tabTab.Rows)
            {
                DataRow x = t.NewRow();
                x["mov_art"] = y["mov_art"];
                x["mov_ard"] = y["MovArd"];
                x["mov_iva"] = y["mov_iva"];
                x["mov_qta"] = y["mov_qta"];
                x["mov_prz"] = y["mov_cos"];
                x["mov_imp"] = y["mov_imp"];

                s = _clsFun.SqlInsertRow(t.TableName, t, x);
                _clsFun.MdbWrite(s, _strConMdb);
            }
        }

        public void SalvaIva(DataTable tabTab)
        {
            string p = "DocIva";
            string s = "DELETE FROM " + p;
            _clsFun.MdbWrite(s, _strConMdb);

            DataTable t = _clsFun.FillTabMdb("DocMov", "SELECT * FROM " + p, true, _strConMdb);

            foreach (DataRow y in tabTab.Rows)
            {
                DataRow x = t.NewRow();
                x["iva_cod"] = y["iva_cod"];
                x["iva_des"] = y["iva_des"];
                x["iva_ali"] = y["iva_ali"];
                x["iva_imp"] = y["iva_imp"];
                x["iva_iva"] = y["iva_iva"];
                x["iva_tot"] = y["iva_tot"];

                s = _clsFun.SqlInsertRow(p, t, x);
                _clsFun.MdbWrite(s, _strConMdb);
            }
        }

    }
}
