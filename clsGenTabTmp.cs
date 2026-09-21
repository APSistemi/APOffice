using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;

namespace APOffice
{
    class clsGenTabTmp
    {
        clsFuncs _clsFun = new clsFuncs();

        public DataTable TabTmpArt(string sTab)
        {
            DataTable t = new DataTable(sTab);
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "tmp_art",
                Caption = "Art",
                MaxLength = 7,
                ReadOnly = false,
                DefaultValue = (String)""
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "tmp_ard",
                Caption = "Descrizione",
                MaxLength = 50,
                ReadOnly = false,
                DefaultValue = (String)""
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "tmp_arb",
                Caption = "Descrizione breve",
                MaxLength = 25,
                ReadOnly = false,
                DefaultValue = (String)""
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "tmp_sta",
                Caption = "Stato",
                MaxLength = 1,
                ReadOnly = false,
                DefaultValue = (String)""
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "tmp_iva",
                Caption = "IVA",
                MaxLength = 3,
                ReadOnly = false,
                DefaultValue = (String)""
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "tmp_umi",
                Caption = "UMI",
                MaxLength = 2,
                ReadOnly = false,
                DefaultValue = (String)""
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "tmp_rep",
                Caption = "Reparto",
                MaxLength = 3,
                ReadOnly = false,
                DefaultValue = (String)""
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "tmp_ecr",
                Caption = "Merceologia",
                MaxLength = 9,
                ReadOnly = false,
                DefaultValue = (String)""
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "tmp_for",
                Caption = "Fornitore",
                MaxLength = 5,
                ReadOnly = false,
                DefaultValue = (String)""
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "tmp_fod",
                Caption = "Fornitore",
                MaxLength = 70,
                ReadOnly = false,
                DefaultValue = (String)""
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "tmp_arf",
                Caption = "Art. fornitore",
                MaxLength = 20,
                ReadOnly = false,
                DefaultValue = (String)""
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "tmp_lis",
                Caption = "Listino",
                MaxLength = 3,
                ReadOnly = false,
                DefaultValue = (String)""
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "tmp_liv",
                Caption = "Listino ivato S/N",
                MaxLength = 1,
                ReadOnly = false,
                DefaultValue = (String)""
            });
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
                DataType = Type.GetType("System.Decimal"),
                ColumnName = "tmp_prp",
                Caption = "Vendita al pubblico",
                ReadOnly = false,
                DefaultValue = (Decimal)0
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.Decimal"),
                ColumnName = "tmp_ppv",
                Caption = "Prezzo vendita pieno in caso di lis=PRO",
                ReadOnly = false,
                DefaultValue = (Decimal)0
            });

            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "tmp_lva",
                Caption = "Listino vendita annullato",
                MaxLength = 3,
                ReadOnly = false,
                DefaultValue = (String)""
            });

            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.Boolean"),
                ColumnName = "tmp_bil",
                Caption = "Bilancia",
                ReadOnly = false,
                DefaultValue = (Boolean)false
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "tmp_tgr",
                Caption = "T.gramm",
                MaxLength = 2,
                ReadOnly = false,
                DefaultValue = (String)""
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.Decimal"),
                ColumnName = "tmp_pxc",
                Caption = "Pezzi",
                ReadOnly = false,
                DefaultValue = (Decimal)0
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.Decimal"),
                ColumnName = "tmp_pne",
                Caption = "Peso netto",
                ReadOnly = false,
                DefaultValue = (Decimal)0
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.Decimal"),
                ColumnName = "tmp_sfr",
                Caption = "Sfrido",
                ReadOnly = false,
                DefaultValue = (Decimal)0
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.Decimal"),
                ColumnName = "tmp_tar",
                Caption = "Tara",
                ReadOnly = false,
                DefaultValue = (Decimal)0
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "tmp_ori",
                Caption = "Origine",
                MaxLength = 3,
                ReadOnly = false,
                DefaultValue = (String)""
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "tmp_cal",
                Caption = "Calibro",
                MaxLength = 3,
                ReadOnly = false,
                DefaultValue = (String)""
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "tmp_cat",
                Caption = "Categoria",
                MaxLength = 3,
                ReadOnly = false,
                DefaultValue = (String)""
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "tmp_reb",
                Caption = "Rep bilancia",
                MaxLength = 1,
                ReadOnly = false,
                DefaultValue = (String)""
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "tmp_gsc",
                Caption = "GG scad",
                MaxLength = 3,
                ReadOnly = false,
                DefaultValue = (String)""
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "tmp_plu",
                Caption = "PLU",
                MaxLength = 6,
                ReadOnly = false,
                DefaultValue = (String)""
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "tmp_tas",
                Caption = "Tasto bilancia",
                MaxLength = 4,
                ReadOnly = false,
                DefaultValue = (String)""
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "tmp_tra",
                Caption = "Tracciabilità",
                MaxLength = 3,
                ReadOnly = false,
                DefaultValue = (String)""
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.Boolean"),
                ColumnName = "tmp_bpz",
                Caption = "A corpo",
                ReadOnly = false,
                DefaultValue = (Boolean)false
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
                DataType = Type.GetType("System.Decimal"),
                ColumnName = "tmp_eaq",
                Caption = "Q.tà x ean",
                ReadOnly = false,
                DefaultValue = (Decimal)0
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.Decimal"),
                ColumnName = "tmp_eap",
                Caption = "Prezzo per ean",
                ReadOnly = false,
                DefaultValue = (Decimal)0
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.Boolean"),
                ColumnName = "tmp_ecp",
                Caption = "Ean con peso",
                ReadOnly = false,
                DefaultValue = (Boolean)false
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.Boolean"),
                ColumnName = "tmp_eaa",
                Caption = "Ean annullato",
                ReadOnly = false,
                DefaultValue = (Boolean)false
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.Decimal"),
                ColumnName = "tmp_epu",
                Caption = "Punti per ean",
                ReadOnly = false,
                DefaultValue = (Decimal)0
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "tmp_cam",
                Caption = "Campagna",
                MaxLength = 3,
                ReadOnly = false,
                DefaultValue = (String)""
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "tmp_eti",
                Caption = "Etichetta",
                MaxLength = 3,
                ReadOnly = false,
                DefaultValue = (String)""
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "tmp_eqp",
                Caption = "Equivalenza prezzo",
                MaxLength = 4,
                ReadOnly = false,
                DefaultValue = (String)""
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "tmp_lot",
                Caption = "Lotto da barcode",
                MaxLength = 15,
                ReadOnly = false,
                DefaultValue = (String)""
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "tmp_img",
                Caption = "Immagine",
                MaxLength = 255,
                ReadOnly = false,
                DefaultValue = (String)""
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.Decimal"),
                ColumnName = "tmp_gia",
                Caption = "Giacenza",
                ReadOnly = false,
                DefaultValue = (Decimal)0
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "tmp_tof",
                Caption = "Tipo offerta",
                MaxLength = 3,
                ReadOnly = false,
                DefaultValue = (String)""
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.Decimal"),
                ColumnName = "tmp_pro",
                Caption = "Prezzo Offerta",
                ReadOnly = false,
                DefaultValue = (Decimal)0
            });

            return t;
        }

        public DataTable TabTmpDivArtApShop(string sTab)
        {
            DataTable t = new DataTable(sTab);
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "tmp_art",
                Caption = "Art",
                MaxLength = 7,
                ReadOnly = false,
                DefaultValue = (String)""
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "tmp_ard",
                Caption = "Descrizione",
                MaxLength = 50,
                ReadOnly = false,
                DefaultValue = (String)""
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "tmp_arb",
                Caption = "Descrizione breve",
                MaxLength = 25,
                ReadOnly = false,
                DefaultValue = (String)""
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "tmp_sta",
                Caption = "Stato",
                MaxLength = 1,
                ReadOnly = false,
                DefaultValue = (String)""
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "tmp_iva",
                Caption = "IVA",
                MaxLength = 3,
                ReadOnly = false,
                DefaultValue = (String)""
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "tmp_umi",
                Caption = "UMI",
                MaxLength = 2,
                ReadOnly = false,
                DefaultValue = (String)""
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "tmp_rep",
                Caption = "Reparto",
                MaxLength = 3,
                ReadOnly = false,
                DefaultValue = (String)""
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "tmp_lis",
                Caption = "Listino",
                MaxLength = 3,
                ReadOnly = false,
                DefaultValue = (String)""
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "tmp_liv",
                Caption = "Listino ivato S/N",
                MaxLength = 1,
                ReadOnly = false,
                DefaultValue = (String)""
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
                DataType = Type.GetType("System.Decimal"),
                ColumnName = "tmp_ppv",
                Caption = "Prezzo pieno",
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
                DataType = Type.GetType("System.Decimal"),
                ColumnName = "tmp_eaq",
                Caption = "Q.tà x ean",
                ReadOnly = false,
                DefaultValue = (Decimal)0
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.Decimal"),
                ColumnName = "tmp_eap",
                Caption = "Prezzo per ean ",
                ReadOnly = false,
                DefaultValue = (Decimal)0
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.Boolean"),
                ColumnName = "tmp_ecp",
                Caption = "Ean con peso",
                ReadOnly = false,
                DefaultValue = (Boolean)false
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.Boolean"),
                ColumnName = "tmp_eab",
                Caption = "Ean bilancia",
                ReadOnly = false,
                DefaultValue = (Boolean)false
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.Boolean"),
                ColumnName = "tmp_eaa",
                Caption = "Ean annullato",
                ReadOnly = false,
                DefaultValue = (Boolean)false
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.Decimal"),
                ColumnName = "tmp_epu",
                Caption = "Prezzo per ean ",
                ReadOnly = false,
                DefaultValue = (Decimal)0
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "tmp_off",
                Caption = "Offerta codice",
                MaxLength = 3,
                ReadOnly = false,
                DefaultValue = (String)""
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "tmp_oft",
                Caption = "Offerta tipo",
                MaxLength = 3,
                ReadOnly = false,
                DefaultValue = (String)""
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.DateTime"),
                ColumnName = "tmp_odi",
                Caption = "Offerta data inizio",
                ReadOnly = false,
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.DateTime"),
                ColumnName = "tmp_odf",
                Caption = "Offerta data fine",
                ReadOnly = false,
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.Decimal"),
                ColumnName = "tmp_ofv",
                Caption = "Offerta val",
                ReadOnly = false,
                DefaultValue = (Decimal)0
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.Decimal"),
                ColumnName = "tmp_ofm",
                Caption = "Offerta Mxn",
                ReadOnly = false,
                DefaultValue = (Decimal)0
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.Decimal"),
                ColumnName = "tmp_ofn",
                Caption = "Offerta mxN",
                ReadOnly = false,
                DefaultValue = (Decimal)0
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.Decimal"),
                ColumnName = "tmp_pun",
                Caption = "Offerta mxN",
                ReadOnly = false,
                DefaultValue = (Decimal)0
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.Decimal"),
                ColumnName = "tmp_mix",
                Caption = "Offerta codice mix",
                ReadOnly = false,
                DefaultValue = (Decimal)0
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "tmp_cam",
                Caption = "Campagna",
                MaxLength = 100,
                ReadOnly = false,
                DefaultValue = (String)""
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "tmp_plu",
                Caption = "PLU",
                MaxLength = 6,
                ReadOnly = false,
                DefaultValue = (String)""
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "tmp_reb",
                Caption = "Reparto bilancia",
                MaxLength = 1,
                ReadOnly = false,
                DefaultValue = (String)""
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.Decimal"),
                ColumnName = "tmp_tar",
                Caption = "Tara",
                ReadOnly = false,
                DefaultValue = (Decimal)0
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.Boolean"),
                ColumnName = "tmp_bpz",
                Caption = "A corpo",
                ReadOnly = false,
                DefaultValue = (Boolean)false
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "tmp_lva",
                Caption = "Listino ventida annullato",
                MaxLength = 3,
                ReadOnly = false,
                DefaultValue = (String)""
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.Boolean"),
                ColumnName = "tmp_cel",
                Caption = "Prodotto celiatico",
                ReadOnly = false,
                DefaultValue = (Boolean)false
            });

            return t;
        }

        public DataTable TabTmpFor(string sTab)
        {
            DataTable t = new DataTable(sTab);

            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "tmp_for",
                Caption = "Fornitore",
                MaxLength = 5,
                ReadOnly = false,
                DefaultValue = (object)""
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "tmp_fod",
                Caption = "Fornitore des",
                MaxLength = 70,
                ReadOnly = false,
                DefaultValue = (object)""
            });

            return t;
        }

        public DataTable TabTmpCli(string sTab)
        {
            DataTable t = new DataTable(sTab);

            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "tmp_cli",
                Caption = "Cliente",
                MaxLength = 5,
                ReadOnly = false,
                DefaultValue = (object)""
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "tmp_cld",
                Caption = "Cliente des",
                MaxLength = 80,
                ReadOnly = false,
                DefaultValue = (object)""
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "tmp_tpa",
                Caption = "Cliente des",
                MaxLength = 3,
                ReadOnly = false,
                DefaultValue = (object)""
            });

            return t;
        }

        public DataTable TabTmpTes(string sTab)
        {
            DataTable t = new DataTable(sTab);

            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "tmp_tes",
                Caption = "Tessera",
                MaxLength = 13,
                ReadOnly = false,
                DefaultValue = (object)""
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "tmp_ted",
                Caption = "Tessera des",
                MaxLength = 50,
                ReadOnly = false,
                DefaultValue = (object)""
            });

            return t;
        }

        public DataTable TabTmpIvaRiep(string sTab)
        {
            DataTable t = new DataTable(sTab);

            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "iva_cod",
                Caption = "Codice",
                MaxLength = 3,
                ReadOnly = false,
                DefaultValue = (object)""
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "iva_neg",
                Caption = "Negozio",
                MaxLength = 3,
                ReadOnly = false,
                DefaultValue = (object)""
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "iva_des",
                Caption = "Des",
                MaxLength = 50,
                ReadOnly = false,
                DefaultValue = (object)""
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.Decimal"),
                ColumnName = "iva_ali",
                Caption = "Des",
                ReadOnly = false,
                DefaultValue = (decimal)0
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.Decimal"),
                ColumnName = "iva_iva",
                Caption = "Iva",
                ReadOnly = false,
                DefaultValue = (decimal)0
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.Decimal"),
                ColumnName = "iva_imp",
                Caption = "Imponibile",
                ReadOnly = false,
                DefaultValue = (decimal)0
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.Decimal"),
                ColumnName = "iva_tot",
                Caption = "Tot",
                ReadOnly = false,
                DefaultValue = (decimal)0
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.Decimal"),
                ColumnName = "iva_arr",
                Caption = "Arr",
                ReadOnly = false,
                DefaultValue = (decimal)0
            });

            return t;
        }

        public DataTable TabTmpPos(string sTab)
        {
            DataTable t = new DataTable(sTab);
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "pos_inv",
                Caption = "Invio",
                MaxLength = 1,
                ReadOnly = false,
                DefaultValue = (String)""
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "pos_tva",
                Caption = "Tipo variazione",
                MaxLength = 3,
                ReadOnly = false,
                DefaultValue = (String)""
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "pos_tvd",
                Caption = "Tipo variazione descriz",
                MaxLength = 50,
                ReadOnly = false,
                DefaultValue = (String)""
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "pos_art",
                Caption = "Art",
                MaxLength = 7,
                ReadOnly = false,
                DefaultValue = (String)""
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "pos_ard",
                Caption = "Descrizione",
                MaxLength = 50,
                ReadOnly = false,
                DefaultValue = (String)""
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "pos_arb",
                Caption = "Descrizione breve",
                MaxLength = 20,
                ReadOnly = false,
                DefaultValue = (String)""
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "pos_sta",
                Caption = "Stato articolo",
                MaxLength = 1,
                ReadOnly = false,
                DefaultValue = (String)""
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "pos_iva",
                Caption = "IVA",
                MaxLength = 3,
                ReadOnly = false,
                DefaultValue = (String)""
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "pos_umi",
                Caption = "UMI",
                MaxLength = 2,
                ReadOnly = false,
                DefaultValue = (String)""
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "pos_tgr",
                Caption = "Tipo gramm",
                MaxLength = 2,
                ReadOnly = false,
                DefaultValue = (String)""
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.Decimal"),
                ColumnName = "pos_pne",
                Caption = "Peso netto",
                ReadOnly = false,
                DefaultValue = (Decimal)0
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.Decimal"),
                ColumnName = "pos_pxc",
                Caption = "Pxc",
                ReadOnly = false,
                DefaultValue = (Decimal)0
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.Decimal"),
                ColumnName = "pos_pvo",
                Caption = "Prezzo Vecchio",
                ReadOnly = false,
                DefaultValue = (Decimal)0
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.Decimal"),
                ColumnName = "pos_mrg",
                Caption = "Margine",
                ReadOnly = false,
                DefaultValue = (Decimal)0
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "pos_rep",
                Caption = "Reparto",
                MaxLength = 3,
                ReadOnly = false,
                DefaultValue = (String)""
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "pos_ecr",
                Caption = "ECR",
                MaxLength = 9,
                ReadOnly = false,
                DefaultValue = (String)""
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "pos_off",
                Caption = "Offerta codice",
                MaxLength = 3,
                ReadOnly = false,
                DefaultValue = (String)""
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "pos_ost",
                Caption = "Offerta stato",
                MaxLength = 1,
                ReadOnly = false,
                DefaultValue = (String)""
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "pos_ofd",
                Caption = "Offerta Descrizione",
                MaxLength = 10,
                ReadOnly = false,
                DefaultValue = (String)""
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.DateTime"),
                ColumnName = "pos_dti",
                Caption = "Inizio offerta",
                ReadOnly = false,
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.DateTime"),
                ColumnName = "pos_dtf",
                Caption = "Fine offerta",
                ReadOnly = false,
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.Decimal"),
                ColumnName = "pos_ova",
                Caption = "Valore",
                ReadOnly = false,
                DefaultValue = (Decimal)0
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.Decimal"),
                ColumnName = "pos_xem",
                Caption = "Emme",
                ReadOnly = false,
                DefaultValue = (Decimal)0
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.Decimal"),
                ColumnName = "pos_xen",
                Caption = "Enne",
                ReadOnly = false,
                DefaultValue = (Decimal)0
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.Decimal"),
                ColumnName = "pos_mix",
                Caption = "Mix match",
                ReadOnly = false,
                DefaultValue = (Decimal)0
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.Decimal"),
                ColumnName = "pos_pun",
                Caption = "Punti",
                ReadOnly = false,
                DefaultValue = (Decimal)0
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "pos_cam",
                Caption = "Campagna",
                MaxLength = 3,
                ReadOnly = false,
                DefaultValue = (String)""
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "pos_neg",
                Caption = "Codice negozio",
                MaxLength = 3,
                ReadOnly = false,
                DefaultValue = (String)""
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.Decimal"),
                ColumnName = "pos_prv",
                Caption = "Vendita",
                ReadOnly = false,
                DefaultValue = (Decimal)0
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "pos_ean",
                Caption = "Barcode",
                MaxLength = 13,
                ReadOnly = false,
                DefaultValue = (String)""
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.Decimal"),
                ColumnName = "pos_eaq",
                Caption = "Q.tà x ean",
                ReadOnly = false,
                DefaultValue = (Decimal)0
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.Decimal"),
                ColumnName = "pos_eap",
                Caption = "Prezzo per ean",
                ReadOnly = false,
                DefaultValue = (Decimal)0
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.Boolean"),
                ColumnName = "pos_ecp",
                Caption = "Ean con peso",
                ReadOnly = false,
                DefaultValue = (Boolean)false
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.Boolean"),
                ColumnName = "pos_eaa",
                Caption = "Barcode annullato",
                ReadOnly = false,
                DefaultValue = (Boolean)false
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.Decimal"),
                ColumnName = "pos_epu",
                Caption = "Punti per ean",
                ReadOnly = false,
                DefaultValue = (Decimal)0
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.DateTime"),
                ColumnName = "pos_day",
                Caption = "Data var",
                ReadOnly = false,
                //DefaultValue = (DateTime)0
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "pos_msg",
                Caption = "Message",
                MaxLength = 100,
                ReadOnly = false,
                DefaultValue = (String)""
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "pos_ori",
                Caption = "Origine variazione",
                MaxLength = 50,
                ReadOnly = false,
                DefaultValue = (String)""
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.Boolean"),
                ColumnName = "bil_bil",
                Caption = "A bilancia",
                ReadOnly = false,
                DefaultValue = (Boolean)false
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "bil_ori",
                Caption = "Origine",
                MaxLength = 3,
                ReadOnly = false,
                DefaultValue = (String)""
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "bil_cal",
                Caption = "Calibro",
                MaxLength = 3,
                ReadOnly = false,
                DefaultValue = (String)""
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "bil_cat",
                Caption = "Categoria",
                MaxLength = 1,
                ReadOnly = false,
                DefaultValue = (String)""
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "bil_reb",
                Caption = "Rep bilancia",
                MaxLength = 1,
                ReadOnly = false,
                DefaultValue = (String)""
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "bil_rbi",
                Caption = "Rep bilancia da tabella reparti",
                MaxLength = 4,
                ReadOnly = false,
                DefaultValue = (String)""
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "bil_gru",
                Caption = "Rep gruppo",
                MaxLength = 2,
                ReadOnly = false,
                DefaultValue = (String)""
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "bil_gsc",
                Caption = "GG scad",
                MaxLength = 3,
                ReadOnly = false,
                DefaultValue = (String)""
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "bil_plu",
                Caption = "PLU",
                MaxLength = 6,
                ReadOnly = false,
                DefaultValue = (String)""
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.Boolean"),
                ColumnName = "bil_bpz",
                Caption = "A corpo",
                ReadOnly = false,
                DefaultValue = (Boolean)false
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "bil_tra",
                Caption = "Tracciabilità",
                MaxLength = 3,
                ReadOnly = false,
                DefaultValue = (String)""
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "bil_tas",
                Caption = "Tasto bilancia",
                MaxLength = 4,
                ReadOnly = false,
                DefaultValue = (String)""
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.Decimal"),
                ColumnName = "bil_tar",
                Caption = "Tara",
                ReadOnly = false,
                DefaultValue = (Decimal)0
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "bil_img",
                Caption = "Immage",
                MaxLength = 255, // Aumentato da 50 a 255
                ReadOnly = false,
                DefaultValue = (String)""
            });

            return t;
        }

        public DataTable TabTmpEti(string sTab)
        {
            DataTable t = new DataTable(sTab);
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "eti_inv",
                Caption = "Invio",
                MaxLength = 1,
                ReadOnly = false,
                DefaultValue = (String)""
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "eti_art",
                Caption = "Art",
                MaxLength = 7,
                ReadOnly = false,
                DefaultValue = (String)""
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "eti_ard",
                Caption = "Descrizione",
                MaxLength = 50,
                ReadOnly = false,
                DefaultValue = (String)""
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "eti_sta",
                Caption = "Stato",
                MaxLength = 1,
                ReadOnly = false,
                DefaultValue = (String)""
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "eti_iva",
                Caption = "IVA",
                MaxLength = 3,
                ReadOnly = false,
                DefaultValue = (String)""
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "eti_ean",
                Caption = "Ean",
                MaxLength = 13,
                ReadOnly = false,
                DefaultValue = (String)""
            });
            //t.Columns.Add(new DataColumn()
            //{
            //    DataType = Type.GetType("System.String"),
            //    ColumnName = "eti_umc",
            //    Caption = "UMI in chiaro",
            //    MaxLength = 2,
            //    ReadOnly = false,
            //    DefaultValue = (String)""
            //});
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.Decimal"),
                ColumnName = "eti_prv",
                Caption = "Vendita",
                ReadOnly = false,
                DefaultValue = (Decimal)0
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "eti_umi",
                Caption = "UMI",
                MaxLength = 2,
                ReadOnly = false,
                DefaultValue = (String)""
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "eti_tgr",
                Caption = "T.gramm.",
                MaxLength = 2,
                ReadOnly = false,
                DefaultValue = (String)""
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.Decimal"),
                ColumnName = "eti_tgv",
                Caption = "T.gramm. valore per kg/L",
                //MaxLength = 2,
                ReadOnly = false,
                DefaultValue = (Decimal)0
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.Decimal"),
                ColumnName = "eti_pxc",
                Caption = "Pxc",
                //MaxLength = 2,
                ReadOnly = false,
                DefaultValue = (Decimal)0
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.Decimal"),
                ColumnName = "eti_pne",
                Caption = "Peso netto",
                ReadOnly = false,
                DefaultValue = (Decimal)0
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "eti_msg",
                Caption = "Message",
                MaxLength = 50,
                ReadOnly = false,
                DefaultValue = (String)""
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "eti_ori",
                Caption = "Origine variazione",
                MaxLength = 50,
                ReadOnly = false,
                DefaultValue = (String)""
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.DateTime"),
                ColumnName = "eti_day",
                Caption = "Data var",
                ReadOnly = false,
                //DefaultValue = (DateTime)0
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "eti_plu",
                Caption = "PLU",
                MaxLength = 6,
                ReadOnly = false,
                DefaultValue = (string)""
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "eti_rep",
                Caption = "Reparto",
                MaxLength = 3,
                ReadOnly = false,
                DefaultValue = (string)""
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "eti_fod",
                Caption = "Fornitore",
                MaxLength = 50,
                ReadOnly = false,
                DefaultValue = (string)""
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "eti_arf",
                Caption = "Art. fornitore",
                MaxLength = 20,
                ReadOnly = false,
                DefaultValue = (string)""
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.Decimal"),
                ColumnName = "eti_pve",
                Caption = "Prezzo pieno",
                ReadOnly = false,
                DefaultValue = (Decimal)0
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.Decimal"),
                ColumnName = "eti_qta",
                Caption = "Numero etichette",
                ReadOnly = false,
                DefaultValue = (Decimal)0
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "eti_cam",
                Caption = "Campagna fidelity",
                MaxLength = 3,
                ReadOnly = false,
                DefaultValue = (string)""
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "eti_off",
                Caption = "Offerta",
                MaxLength = 3,
                ReadOnly = false,
                DefaultValue = (string)""
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "eti_oft",
                Caption = "Tipo offerta",
                MaxLength = 3,
                ReadOnly = false,
                DefaultValue = (string)""
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.Decimal"),
                ColumnName = "eti_xem",
                Caption = "Emme",
                ReadOnly = false,
                DefaultValue = (Decimal)0
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.Decimal"),
                ColumnName = "eti_xen",
                Caption = "Enne",
                ReadOnly = false,
                DefaultValue = (Decimal)0
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.DateTime"),
                ColumnName = "eti_odi",
                Caption = "Data inizio offerta",
                ReadOnly = false
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.DateTime"),
                ColumnName = "eti_odf",
                Caption = "Data fine offerta",
                ReadOnly = false
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "eti_bor",
                Caption = "Bilancia origine",
                MaxLength = 20,
                ReadOnly = false,
                DefaultValue = (string)""
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "eti_bcl",
                Caption = "Bilancia calibro",
                MaxLength = 20,
                ReadOnly = false,
                DefaultValue = (string)""
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "eti_bct",
                Caption = "Bilancia categoria",
                MaxLength = 20,
                ReadOnly = false,
                DefaultValue = (string)""
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "eti_eti",
                Caption = "Tipo etichetta",
                MaxLength = 3,
                ReadOnly = false,
                DefaultValue = (string)""
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "eti_ecr",
                Caption = "ECR",
                MaxLength = 9,
                ReadOnly = false,
                DefaultValue = (string)""
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "eti_tas",
                Caption = "Tasto",
                MaxLength = 4,
                ReadOnly = false,
                DefaultValue = (string)""
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "eti_img",
                Caption = "Immagine articolo",
                MaxLength = 255, // Aumentato da 100 a 255 per supportare percorsi completi aggiornati
                ReadOnly = false,
                DefaultValue = (string)""
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "eti_red",
                Caption = "Reparto descrizione",
                MaxLength = 50,
                ReadOnly = false,
                DefaultValue = (string)""
            });

            return t;
        }

        public DataTable TabTmpForDiv(string sTab)
        {
            DataTable t = new DataTable(sTab);
            if (sTab == "ForDivTes")
            {
                t.Columns.Add(new DataColumn()
                {
                    DataType = Type.GetType("System.String"),
                    ColumnName = "dit_cho",
                    Caption = "Tipo variazione",
                    MaxLength = 1,
                    ReadOnly = false,
                    DefaultValue = (String)"S"
                });
                t.Columns.Add(new DataColumn()
                {
                    DataType = Type.GetType("System.String"),
                    ColumnName = "dit_tva",
                    Caption = "Tipo variazione",
                    MaxLength = 3,
                    ReadOnly = false,
                    DefaultValue = (String)""
                });
                t.Columns.Add(new DataColumn()
                {
                    DataType = Type.GetType("System.String"),
                    ColumnName = "dit_neg",
                    Caption = "Negozio",
                    MaxLength = 3,
                    ReadOnly = false,
                    DefaultValue = (String)""
                });
                t.Columns.Add(new DataColumn()
                {
                    DataType = Type.GetType("System.String"),
                    ColumnName = "dit_num",
                    Caption = "Numero",
                    MaxLength = 3,
                    ReadOnly = false,
                    DefaultValue = (String)""
                });
                t.Columns.Add(new DataColumn()
                {
                    DataType = Type.GetType("System.String"),
                    ColumnName = "dit_for",
                    Caption = "For",
                    MaxLength = 5,
                    ReadOnly = false,
                    DefaultValue = (String)""
                });
                t.Columns.Add(new DataColumn()
                {
                    DataType = Type.GetType("System.String"),
                    ColumnName = "dit_fil",
                    Caption = "File",
                    MaxLength = 150,
                    ReadOnly = false,
                    DefaultValue = (String)""
                });
                t.Columns.Add(new DataColumn()
                {
                    DataType = Type.GetType("System.String"),
                    ColumnName = "dit_des",
                    Caption = "Descrizione",
                    MaxLength = 50,
                    ReadOnly = false,
                    DefaultValue = (String)""
                });
                t.Columns.Add(new DataColumn()
                {
                    DataType = Type.GetType("System.DateTime"),
                    ColumnName = "dit_sii",
                    Caption = "Data sell-in inizio",
                    ReadOnly = false,
                });
                t.Columns.Add(new DataColumn()
                {
                    DataType = Type.GetType("System.DateTime"),
                    ColumnName = "dit_sif",
                    Caption = "Data sell-in fine",
                    ReadOnly = false,
                });
                t.Columns.Add(new DataColumn()
                {
                    DataType = Type.GetType("System.DateTime"),
                    ColumnName = "dit_soi",
                    Caption = "Data sell-in inizio",
                    ReadOnly = false,
                });
                t.Columns.Add(new DataColumn()
                {
                    DataType = Type.GetType("System.DateTime"),
                    ColumnName = "dit_sof",
                    Caption = "Data sell-in fine",
                    ReadOnly = false,
                });
                t.Columns.Add(new DataColumn()
                {
                    DataType = Type.GetType("System.String"),
                    ColumnName = "dit_ndo",
                    Caption = "Numero documento",
                    MaxLength = 10,
                    ReadOnly = false,
                    DefaultValue = (String)""
                });
                t.Columns.Add(new DataColumn()
                {
                    DataType = Type.GetType("System.DateTime"),
                    ColumnName = "dit_ddo",
                    Caption = "Data documento",
                    ReadOnly = false,
                });
                t.Columns.Add(new DataColumn()
                {
                    DataType = Type.GetType("System.String"),
                    ColumnName = "dit_pth",
                    Caption = "Path file",
                    MaxLength = 100,
                    ReadOnly = false,
                    DefaultValue = (String)""
                });
                t.Columns.Add(new DataColumn()
                {
                    DataType = Type.GetType("System.Decimal"),
                    ColumnName = "dit_nri",
                    Caption = "Numero righe",
                    ReadOnly = false,
                    DefaultValue = (Decimal)0
                });
                t.Columns.Add(new DataColumn()
                {
                    DataType = Type.GetType("System.String"),
                    ColumnName = "dit_nva",
                    Caption = "Numero variazioni",
                    MaxLength = 10,
                    ReadOnly = false,
                    DefaultValue = (String)""
                });
                t.Columns.Add(new DataColumn()
                {
                    DataType = Type.GetType("System.String"),
                    ColumnName = "dit_tfo",
                    Caption = "Codice tabella variazioni",
                    MaxLength = 10,
                    ReadOnly = false,
                    DefaultValue = (String)""
                });
            }
            if (sTab == "ForDivRig")
            {
                t.Columns.Add(new DataColumn()
                {
                    DataType = Type.GetType("System.String"),
                    ColumnName = "div_for",
                    Caption = "Fornitore",
                    MaxLength = 5,
                    ReadOnly = false,
                    DefaultValue = (String)""
                });
                t.Columns.Add(new DataColumn()
                {
                    DataType = Type.GetType("System.String"),
                    ColumnName = "div_tva",
                    Caption = "Tipo",
                    MaxLength = 3,
                    ReadOnly = false,
                    DefaultValue = (String)""
                });
                t.Columns.Add(new DataColumn()
                {
                    DataType = Type.GetType("System.DateTime"),
                    ColumnName = "div_dva",
                    Caption = "Data variaz",
                    ReadOnly = false
                });
                t.Columns.Add(new DataColumn()
                {
                    DataType = Type.GetType("System.DateTime"),
                    ColumnName = "div_sif",
                    Caption = "Data fine per offerte in acq",
                    ReadOnly = false
                });
                t.Columns.Add(new DataColumn()
                {
                    DataType = Type.GetType("System.String"),
                    ColumnName = "div_num",
                    Caption = "Numero",
                    MaxLength = 3,
                    ReadOnly = false,
                    DefaultValue = (String)""
                });
                t.Columns.Add(new DataColumn()
                {
                    DataType = Type.GetType("System.String"),
                    ColumnName = "div_tip",
                    Caption = "Tipo variazione/offerta",
                    MaxLength = 3,
                    ReadOnly = false,
                    DefaultValue = (String)""
                });
                t.Columns.Add(new DataColumn()
                {
                    DataType = Type.GetType("System.String"),
                    ColumnName = "div_nri",
                    Caption = "Numero riga",
                    MaxLength = 5,
                    ReadOnly = false,
                    DefaultValue = (String)""
                });
                t.Columns.Add(new DataColumn()
                {
                    DataType = Type.GetType("System.String"),
                    ColumnName = "div_art",
                    Caption = "Articolo",
                    MaxLength = 20,
                    ReadOnly = false,
                    DefaultValue = (String)""
                });
                t.Columns.Add(new DataColumn()
                {
                    DataType = Type.GetType("System.String"),
                    ColumnName = "div_arf",
                    Caption = "Articolo",
                    MaxLength = 15,
                    ReadOnly = false,
                    DefaultValue = (String)""
                });
                t.Columns.Add(new DataColumn()
                {
                    DataType = Type.GetType("System.String"),
                    ColumnName = "div_ard",
                    Caption = "Descrizione",
                    MaxLength = 50,
                    ReadOnly = false,
                    DefaultValue = (String)""
                });
                t.Columns.Add(new DataColumn()
                {
                    DataType = Type.GetType("System.String"),
                    ColumnName = "div_adb",
                    Caption = "Descrizione breve",
                    MaxLength = 25,
                    ReadOnly = false,
                    DefaultValue = (String)""
                });
                t.Columns.Add(new DataColumn()
                {
                    DataType = Type.GetType("System.String"),
                    ColumnName = "for_tgr",
                    Caption = "Tipo grammatura",
                    MaxLength = 2,
                    ReadOnly = false,
                    DefaultValue = (String)""
                });
                t.Columns.Add(new DataColumn()
                {
                    DataType = Type.GetType("System.String"),
                    ColumnName = "div_tgr",
                    Caption = "Tipo grammatura",
                    MaxLength = 2,
                    ReadOnly = false,
                    DefaultValue = (String)""
                });
                t.Columns.Add(new DataColumn()
                {
                    DataType = Type.GetType("System.String"),
                    ColumnName = "for_umi",
                    Caption = "UM",
                    MaxLength = 2,
                    ReadOnly = false,
                    DefaultValue = (String)""
                });
                t.Columns.Add(new DataColumn()
                {
                    DataType = Type.GetType("System.String"),
                    ColumnName = "div_umi",
                    Caption = "UM",
                    MaxLength = 2,
                    ReadOnly = false,
                    DefaultValue = (String)""
                });
                t.Columns.Add(new DataColumn()
                {
                    DataType = Type.GetType("System.String"),
                    ColumnName = "for_rep",
                    Caption = "Reparto",
                    MaxLength = 3,
                    ReadOnly = false,
                    DefaultValue = (String)""
                });
                t.Columns.Add(new DataColumn()
                {
                    DataType = Type.GetType("System.String"),
                    ColumnName = "div_rep",
                    Caption = "Reparto",
                    MaxLength = 3,
                    ReadOnly = false,
                    DefaultValue = (String)""
                });
                t.Columns.Add(new DataColumn()
                {
                    DataType = Type.GetType("System.String"),
                    ColumnName = "for_ecr",
                    Caption = "ECR",
                    MaxLength = 12,
                    ReadOnly = false,
                    DefaultValue = (String)""
                });
                t.Columns.Add(new DataColumn()
                {
                    DataType = Type.GetType("System.String"),
                    ColumnName = "div_ec1",
                    Caption = "ECR",
                    MaxLength = 3,
                    ReadOnly = false,
                    DefaultValue = (String)""
                });
                t.Columns.Add(new DataColumn()
                {
                    DataType = Type.GetType("System.String"),
                    ColumnName = "div_ec2",
                    Caption = "ECR",
                    MaxLength = 3,
                    ReadOnly = false,
                    DefaultValue = (String)""
                });
                t.Columns.Add(new DataColumn()
                {
                    DataType = Type.GetType("System.String"),
                    ColumnName = "div_ec3",
                    Caption = "ECR",
                    MaxLength = 3,
                    ReadOnly = false,
                    DefaultValue = (String)""
                });
                t.Columns.Add(new DataColumn()
                {
                    DataType = Type.GetType("System.String"),
                    ColumnName = "for_iva",
                    Caption = "IVA",
                    MaxLength = 3,
                    ReadOnly = false,
                    DefaultValue = (String)""
                });
                t.Columns.Add(new DataColumn()
                {
                    DataType = Type.GetType("System.String"),
                    ColumnName = "div_iva",
                    Caption = "IVA",
                    MaxLength = 3,
                    ReadOnly = false,
                    DefaultValue = (String)""
                });
                t.Columns.Add(new DataColumn()
                {
                    DataType = Type.GetType("System.Decimal"),
                    ColumnName = "div_pxc",
                    Caption = "PxC",
                    ReadOnly = false,
                    DefaultValue = (Decimal)0
                });
                t.Columns.Add(new DataColumn()
                {
                    DataType = Type.GetType("System.Decimal"),
                    ColumnName = "div_pne",
                    Caption = "Grammat.",
                    ReadOnly = false,
                    DefaultValue = (Decimal)0
                });
                t.Columns.Add(new DataColumn()
                {
                    DataType = Type.GetType("System.String"),
                    ColumnName = "div_tpd",
                    Caption = "Tipo doc",
                    MaxLength = 1,
                    ReadOnly = false,
                    DefaultValue = (String)""
                });
                t.Columns.Add(new DataColumn()
                {
                    DataType = Type.GetType("System.Decimal"),
                    ColumnName = "div_qta",
                    Caption = "Qtà in doc",
                    ReadOnly = false,
                    DefaultValue = (Decimal)0
                });
                //t.Columns.Add(new DataColumn()
                //{
                //    DataType = Type.GetType("System.Decimal"),
                //    ColumnName = "div_cof",
                //    Caption = "Costo fattura",
                //    ReadOnly = false,
                //    DefaultValue = (Decimal)0
                //});
                t.Columns.Add(new DataColumn()
                {
                    DataType = Type.GetType("System.Decimal"),
                    ColumnName = "div_tof",
                    Caption = "Importo riga fattura",
                    ReadOnly = false,
                    DefaultValue = (Decimal)0
                });
                t.Columns.Add(new DataColumn()
                {
                    DataType = Type.GetType("System.String"),
                    ColumnName = "div_ean",
                    Caption = "Barcode",
                    MaxLength = 13,
                    ReadOnly = false,
                    DefaultValue = (String)""
                });
                t.Columns.Add(new DataColumn()
                {
                    DataType = Type.GetType("System.String"),
                    ColumnName = "div_tri",
                    Caption = "Tipo riga",
                    MaxLength = 1,
                    ReadOnly = false,
                    DefaultValue = (String)""
                });
                t.Columns.Add(new DataColumn()
                {
                    DataType = Type.GetType("System.Decimal"),
                    ColumnName = "div_cof",
                    Caption = "Costo fattura",
                    ReadOnly = false,
                    DefaultValue = (Decimal)0
                });
                t.Columns.Add(new DataColumn()
                {
                    DataType = Type.GetType("System.Decimal"),
                    ColumnName = "art_cos",
                    Caption = "Costo listino",
                    ReadOnly = false,
                    DefaultValue = (Decimal)0
                });
                t.Columns.Add(new DataColumn()
                {
                    DataType = Type.GetType("System.Decimal"),
                    ColumnName = "div_cos",
                    Caption = "Costo listino ????",
                    ReadOnly = false,
                    DefaultValue = (Decimal)0
                });
                t.Columns.Add(new DataColumn()
                {
                    DataType = Type.GetType("System.Decimal"),
                    ColumnName = "dif_cos",
                    Caption = "Differenza listino",
                    ReadOnly = false,
                    DefaultValue = (Decimal)0
                });
                t.Columns.Add(new DataColumn()
                {
                    DataType = Type.GetType("System.Decimal"),
                    ColumnName = "div_imp",
                    Caption = "Importo ???",
                    ReadOnly = false,
                    DefaultValue = (Decimal)0
                });
                t.Columns.Add(new DataColumn()
                {
                    DataType = Type.GetType("System.String"),
                    ColumnName = "liv_lis",
                    Caption = "Listino",
                    MaxLength = 3,
                    ReadOnly = false,
                    DefaultValue = (String)""
                });
                t.Columns.Add(new DataColumn()
                {
                    DataType = Type.GetType("System.Decimal"),
                    ColumnName = "art_sfr",
                    Caption = "Sfrido",
                    ReadOnly = false,
                    DefaultValue = (Decimal)0
                });
                t.Columns.Add(new DataColumn()
                {
                    DataType = Type.GetType("System.Decimal"),
                    ColumnName = "art_prp",
                    Caption = "Prezzo pubblico consigliato",
                    ReadOnly = false,
                    DefaultValue = (Decimal)0
                });
                t.Columns.Add(new DataColumn()
                {
                    DataType = Type.GetType("System.Decimal"),
                    ColumnName = "art_mav",
                    Caption = "Margine",
                    ReadOnly = false,
                    DefaultValue = (Decimal)0
                });
                t.Columns.Add(new DataColumn()
                {
                    DataType = Type.GetType("System.Decimal"),
                    ColumnName = "div_prp",
                    Caption = "Prezzo pubblico consigliato",
                    ReadOnly = false,
                    DefaultValue = (Decimal)0
                });
                t.Columns.Add(new DataColumn()
                {
                    DataType = Type.GetType("System.Decimal"),
                    ColumnName = "div_mav",
                    Caption = "Margine",
                    ReadOnly = false,
                    DefaultValue = (Decimal)0
                });
                t.Columns.Add(new DataColumn()
                {
                    DataType = Type.GetType("System.Decimal"),
                    ColumnName = "var_prp",
                    Caption = "Prezzo variato",
                    ReadOnly = false,
                    DefaultValue = (Decimal)0
                });
                t.Columns.Add(new DataColumn()
                {
                    DataType = Type.GetType("System.Decimal"),
                    ColumnName = "var_mav",
                    Caption = "Margine",
                    ReadOnly = false,
                    DefaultValue = (Decimal)0
                });
                t.Columns.Add(new DataColumn()
                {
                    DataType = Type.GetType("System.String"),
                    ColumnName = "div_sta",
                    Caption = "Stato",
                    MaxLength = 1,
                    ReadOnly = false,
                    DefaultValue = (String)""
                });
                t.Columns.Add(new DataColumn()
                {
                    DataType = Type.GetType("System.String"),
                    ColumnName = "div_stf",
                    Caption = "Stato fornitore",
                    MaxLength = 1,
                    ReadOnly = false,
                    DefaultValue = (String)""
                });
                t.Columns.Add(new DataColumn()
                {
                    DataType = Type.GetType("System.String"),
                    ColumnName = "div_fo2",
                    Caption = "Fornitore consegnatario",
                    MaxLength = 10,
                    ReadOnly = false,
                    DefaultValue = (String)""
                });
                t.Columns.Add(new DataColumn()
                {
                    DataType = Type.GetType("System.String"),
                    ColumnName = "div_af2",
                    Caption = "Codice del consegnatario",
                    MaxLength = 20,
                    ReadOnly = false,
                    DefaultValue = (String)""
                });
                t.Columns.Add(new DataColumn()
                {
                    DataType = Type.GetType("System.String"),
                    ColumnName = "div_off",
                    Caption = "Offerta",
                    MaxLength = 50,
                    ReadOnly = false,
                    DefaultValue = (String)""
                });
                t.Columns.Add(new DataColumn()
                {
                    DataType = Type.GetType("System.String"),
                    ColumnName = "off_tip",
                    Caption = "tipo Offerta",
                    MaxLength = 10,
                    ReadOnly = false,
                    DefaultValue = (String)""
                });
                t.Columns.Add(new DataColumn()
                {
                    DataType = Type.GetType("System.Boolean"),
                    ColumnName = "off_ann",
                    Caption = "Annullo",
                    ReadOnly = false,
                    DefaultValue = (Boolean)false
                });
                t.Columns.Add(new DataColumn()
                {
                    DataType = Type.GetType("System.String"),
                    ColumnName = "off_mxn",
                    Caption = "Offerta MxN",
                    MaxLength = 3,
                    ReadOnly = false,
                    DefaultValue = (String)""
                });
                t.Columns.Add(new DataColumn()
                {
                    DataType = Type.GetType("System.String"),
                    ColumnName = "div_eti",
                    Caption = "Etichetta",
                    MaxLength = 1,
                    ReadOnly = false,
                    DefaultValue = (String)""
                });
                t.Columns.Add(new DataColumn()
                {
                    DataType = Type.GetType("System.String"),
                    ColumnName = "div_bil",
                    Caption = "Art bilancia",
                    MaxLength = 1,
                    ReadOnly = false,
                    DefaultValue = (String)""
                });
                t.Columns.Add(new DataColumn()
                {
                    DataType = Type.GetType("System.String"),
                    ColumnName = "div_plu",
                    Caption = "PLU",
                    MaxLength = 6,
                    ReadOnly = false,
                    DefaultValue = (String)""
                });
                t.Columns.Add(new DataColumn()
                {
                    DataType = Type.GetType("System.String"),
                    ColumnName = "div_reb",
                    Caption = "Reparto bilancia",
                    MaxLength = 1,
                    ReadOnly = false,
                    DefaultValue = (String)""
                });
                t.Columns.Add(new DataColumn()
                {
                    DataType = Type.GetType("System.String"),
                    ColumnName = "div_ori",
                    Caption = "Origine",
                    MaxLength = 3,
                    ReadOnly = false,
                    DefaultValue = (String)""
                });
                t.Columns.Add(new DataColumn()
                {
                    DataType = Type.GetType("System.String"),
                    ColumnName = "div_cal",
                    Caption = "Calibro",
                    MaxLength = 3,
                    ReadOnly = false,
                    DefaultValue = (String)""
                });
                t.Columns.Add(new DataColumn()
                {
                    DataType = Type.GetType("System.String"),
                    ColumnName = "div_cat",
                    Caption = "Categoria",
                    MaxLength = 3,
                    ReadOnly = false,
                    DefaultValue = (String)""
                });


                t.Columns.Add(new DataColumn()
                {
                    DataType = Type.GetType("System.String"),
                    ColumnName = "for_ori",
                    Caption = "Origine",
                    MaxLength = 20,
                    ReadOnly = false,
                    DefaultValue = (String)""
                });
                t.Columns.Add(new DataColumn()
                {
                    DataType = Type.GetType("System.String"),
                    ColumnName = "for_cal",
                    Caption = "Calibro",
                    MaxLength = 20,
                    ReadOnly = false,
                    DefaultValue = (String)""
                });
                t.Columns.Add(new DataColumn()
                {
                    DataType = Type.GetType("System.String"),
                    ColumnName = "for_cat",
                    Caption = "Categoria",
                    MaxLength = 20,
                    ReadOnly = false,
                    DefaultValue = (String)""
                });

                t.Columns.Add(new DataColumn()
                {
                    DataType = Type.GetType("System.String"),
                    ColumnName = "div_gsc",
                    Caption = "Giorni conservazione",
                    MaxLength = 3,
                    ReadOnly = false,
                    DefaultValue = (String)""
                });
                t.Columns.Add(new DataColumn()
                {
                    DataType = Type.GetType("System.String"),
                    ColumnName = "div_bpz",
                    Caption = "A pz in bilancia",
                    MaxLength = 1,
                    ReadOnly = false,
                    DefaultValue = (String)""
                });
                t.Columns.Add(new DataColumn()
                {
                    DataType = Type.GetType("System.String"),
                    ColumnName = "div_var",
                    Caption = "Variazione",
                    MaxLength = 1,
                    ReadOnly = false,
                    DefaultValue = (String)""
                });
                t.Columns.Add(new DataColumn()
                {
                    DataType = Type.GetType("System.String"),
                    ColumnName = "div_ing",
                    Caption = "Ingredienti",
                    MaxLength = 200,
                    ReadOnly = false,
                    DefaultValue = (String)""
                });
            }
            if (sTab == "ForDivEan")
            {
                t.Columns.Add(new DataColumn()
                {
                    DataType = Type.GetType("System.String"),
                    ColumnName = "die_for",
                    Caption = "For",
                    MaxLength = 5,
                    ReadOnly = false,
                    DefaultValue = (String)""
                });
                t.Columns.Add(new DataColumn()
                {
                    DataType = Type.GetType("System.String"),
                    ColumnName = "die_arf",
                    Caption = "Arf",
                    MaxLength = 15,
                    ReadOnly = false,
                    DefaultValue = (String)""
                });
                t.Columns.Add(new DataColumn()
                {
                    DataType = Type.GetType("System.String"),
                    ColumnName = "die_ean",
                    Caption = "Ean",
                    MaxLength = 13,
                    ReadOnly = false,
                    DefaultValue = (String)""
                });
                t.Columns.Add(new DataColumn()
                {
                    DataType = System.Type.GetType("System.Decimal"),
                    ColumnName = "die_qta",
                    Caption = "Qta",
                    ReadOnly = false,
                    DefaultValue = (object)0
                });
                t.Columns.Add(new DataColumn()
                {
                    DataType = System.Type.GetType("System.Decimal"),
                    ColumnName = "die_prv",
                    Caption = "Prv",
                    ReadOnly = false,
                    DefaultValue = (object)0
                });
                t.Columns.Add(new DataColumn()
                {
                    DataType = System.Type.GetType("System.Boolean"),
                    ColumnName = "die_ecp",
                    Caption = "Ean con peso",
                    ReadOnly = false,
                    DefaultValue = (Boolean)false
                });
                t.Columns.Add(new DataColumn()
                {
                    DataType = System.Type.GetType("System.Boolean"),
                    ColumnName = "die_bil",
                    Caption = "Per bilancia",
                    ReadOnly = false,
                    DefaultValue = (Boolean)false
                });
                t.Columns.Add(new DataColumn()
                {
                    DataType = System.Type.GetType("System.Boolean"),
                    ColumnName = "die_ann",
                    Caption = "Annullato",
                    ReadOnly = false,
                    DefaultValue = (Boolean)false
                });

            }
            return t;
        }

        public DataTable TabTmpVet(string sTab)
        {
            DataTable t = new DataTable(sTab);
            t.Columns.Add(new DataColumn()
            {
                DataType = System.Type.GetType("System.String"),
                ColumnName = "VenNeg",
                Caption = "Neg",
                MaxLength = 6,
                ReadOnly = false,
                DefaultValue = (object)""
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = System.Type.GetType("System.DateTime"),
                ColumnName = "VenDay",
                Caption = "Day",
                ReadOnly = false
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = System.Type.GetType("System.String"),
                ColumnName = "VenOra",
                Caption = "Ora",
                MaxLength = 4,
                ReadOnly = false,
                DefaultValue = (object)""
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = System.Type.GetType("System.String"),
                ColumnName = "VenPos",
                Caption = "Pos",
                MaxLength = 2,
                ReadOnly = false,
                DefaultValue = (object)""
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = System.Type.GetType("System.String"),
                ColumnName = "VenSco",
                Caption = "Sco",
                MaxLength = 5,
                ReadOnly = false,
                DefaultValue = (object)""
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = System.Type.GetType("System.String"),
                ColumnName = "VenFid",
                Caption = "Fid",
                MaxLength = 13,
                ReadOnly = false,
                DefaultValue = (object)""
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = System.Type.GetType("System.Decimal"),
                ColumnName = "VenImp",
                Caption = "Imp",
                ReadOnly = false,
                DefaultValue = (object)0
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = System.Type.GetType("System.Decimal"),
                ColumnName = "VenScp",
                Caption = "Sconto",
                ReadOnly = false,
                DefaultValue = (object)0
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = System.Type.GetType("System.Decimal"),
                ColumnName = "VenSct",
                Caption = "Sconto",
                ReadOnly = false,
                DefaultValue = (object)0
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = System.Type.GetType("System.Decimal"),
                ColumnName = "VenPun",
                Caption = "Punti",
                ReadOnly = false,
                DefaultValue = (object)0
            });

            return t;
        }

        public DataTable TabTmpVenduto(string sTab)
        {
            return new DataTable(sTab)
            {
                Columns = {
                new DataColumn()
                {
                  DataType = Type.GetType("System.String"),
                  ColumnName = "ven_tip",
                  Caption = "Tipo",
                  MaxLength = 3,
                  ReadOnly = false
                },
                new DataColumn()
                {
                  DataType = Type.GetType("System.DateTime"),
                  ColumnName = "ven_day",
                  Caption = "Day",
                  ReadOnly = false
                },
                new DataColumn()
                {
                  DataType = Type.GetType("System.String"),
                  ColumnName = "ven_neg",
                  Caption = "Neg",
                  MaxLength = 6,
                  ReadOnly = false
                },
                new DataColumn()
                {
                  DataType = Type.GetType("System.String"),
                  ColumnName = "ven_des",
                  Caption = "Neg",
                  MaxLength = 50,
                  ReadOnly = false
                },
                new DataColumn()
                {
                  DataType = Type.GetType("System.Decimal"),
                  ColumnName = "ven_imp",
                  Caption = "Imp",
                  ReadOnly = false
                }
              }
            };
        }

        public DataTable TabTmpVenReparto(string sTab)
        {
            return new DataTable(sTab)
            {
                Columns = {
                new DataColumn()
                {
                  DataType = Type.GetType("System.String"),
                  ColumnName = "vre_neg",
                  Caption = "Neg",
                  MaxLength = 3,
                  ReadOnly = false
                },
                new DataColumn()
                {
                  DataType = Type.GetType("System.DateTime"),
                  ColumnName = "vre_day",
                  Caption = "Day",
                  ReadOnly = false
                },
                new DataColumn()
                {
                  DataType = Type.GetType("System.String"),
                  ColumnName = "vre_rep",
                  Caption = "Reparto",
                  MaxLength = 3,
                  ReadOnly = false
                },
                new DataColumn()
                {
                  DataType = Type.GetType("System.String"),
                  ColumnName = "vre_red",
                  Caption = "Neg",
                  MaxLength = 50,
                  ReadOnly = false
                },
                new DataColumn()
                {
                  DataType = Type.GetType("System.Decimal"),
                  ColumnName = "vre_ven",
                  Caption = "Imp",
                  ReadOnly = false
                }
              }
            };
        }

        public DataTable TabTmpMsg(string sTab)
        {
            return new DataTable(sTab)
            {
                Columns = {
                new DataColumn()
                {
                  DataType = Type.GetType("System.String"),
                  ColumnName = "msg_co1",
                  Caption = "Codice 1",
                  MaxLength = 20,
                  ReadOnly = false
                },
                new DataColumn()
                {
                  DataType = Type.GetType("System.String"),
                  ColumnName = "msg_co2",
                  Caption = "Codice 2",
                  MaxLength = 20,
                  ReadOnly = false
                },
                new DataColumn()
                {
                  DataType = Type.GetType("System.String"),
                  ColumnName = "msg_des",
                  Caption = "Codice 2",
                  MaxLength = 100,
                  ReadOnly = false
                },
                new DataColumn()
                {
                  DataType = Type.GetType("System.String"),
                  ColumnName = "msg_msg",
                  Caption = "Messaggio",
                  MaxLength = 100,
                  ReadOnly = false
                }
            }
            };
        }

        public DataTable TabTmpDocMov(string sTab)
        {
            DataTable t = new DataTable(sTab);

            if (sTab == "DocTes")
            {
                t = new DataTable(sTab)
                {
                    Columns = {
                    new DataColumn()
                    {
                      DataType = Type.GetType("System.String"),
                      ColumnName = "tmp_cfo",
                      Caption = "Codice fornitore",
                      MaxLength = 5,
                      ReadOnly = false
                    },
                    new DataColumn()
                    {
                      DataType = Type.GetType("System.String"),
                      ColumnName = "tmp_rag",
                      Caption = "Ragione sociale",
                      MaxLength = 100,
                      ReadOnly = false
                    },
                    new DataColumn()
                    {
                      DataType = Type.GetType("System.String"),
                      ColumnName = "tmp_ra2",
                      Caption = "Ragione sociale",
                      MaxLength = 100,
                      ReadOnly = false
                    },
                    new DataColumn()
                    {
                      DataType = Type.GetType("System.String"),
                      ColumnName = "tmp_piv",
                      Caption = "P.Iva",
                      MaxLength = 11,
                      ReadOnly = false
                    },
                    new DataColumn()
                    {
                      DataType = Type.GetType("System.String"),
                      ColumnName = "tmp_cfi",
                      Caption = "Codice fiscale",
                      MaxLength = 16,
                      ReadOnly = false
                    },
                    new DataColumn()
                    {
                      DataType = Type.GetType("System.String"),
                      ColumnName = "tmp_ind",
                      Caption = "Indirizzo",
                      MaxLength = 50,
                      ReadOnly = false
                    },
                    new DataColumn()
                    {
                      DataType = Type.GetType("System.String"),
                      ColumnName = "tmp_ncv",
                      Caption = "Numero civico",
                      MaxLength = 5,
                      ReadOnly = false
                    },
                    new DataColumn()
                    {
                      DataType = Type.GetType("System.String"),
                      ColumnName = "tmp_loc",
                      Caption = "Località",
                      MaxLength = 50,
                      ReadOnly = false
                    },
                    new DataColumn()
                    {
                      DataType = Type.GetType("System.String"),
                      ColumnName = "tmp_ndo",
                      Caption = "N.documento",
                      MaxLength = 20,
                      ReadOnly = false
                    },
                    new DataColumn()
                    {
                      DataType = Type.GetType("System.DateTime"),
                      ColumnName = "tmp_ddo",
                      Caption = "Data documento",
                      ReadOnly = false
                    },
                    new DataColumn()
                    {
                      DataType = Type.GetType("System.String"),
                      ColumnName = "tmp_tpg",
                      Caption = "Tipo pagamento",
                      MaxLength = 50,
                      ReadOnly = false
                    },
                  }
                };
            }
            if (sTab == "DocAll")
            {
                t = new DataTable(sTab)
                {
                    Columns = {
                    new DataColumn()
                    {
                      DataType = Type.GetType("System.String"),
                      ColumnName = "tmp_des",
                      Caption = "Descrizione",
                      MaxLength = 50,
                      ReadOnly = false
                    },
                    new DataColumn()
                    {
                      DataType = Type.GetType("System.String"),
                      ColumnName = "tmp_neg",
                      Caption = "Negozio",
                      MaxLength = 3,
                      ReadOnly = false
                    },
                    new DataColumn()
                    {
                      DataType = Type.GetType("System.String"),
                      ColumnName = "tmp_cfo",
                      Caption = "Codice cliente",
                      MaxLength = 5,
                      ReadOnly = false
                    },
                    new DataColumn()
                    {
                      DataType = Type.GetType("System.String"),
                      ColumnName = "tmp_rag",
                      Caption = "Ragione sociale",
                      MaxLength = 100,
                      ReadOnly = false
                    },
                    new DataColumn()
                    {
                      DataType = Type.GetType("System.String"),
                      ColumnName = "tmp_piv",
                      Caption = "P.Iva",
                      MaxLength = 11,
                      ReadOnly = false
                    },
                    new DataColumn()
                    {
                      DataType = Type.GetType("System.String"),
                      ColumnName = "tmp_cfi",
                      Caption = "Codice fiscale",
                      MaxLength = 16,
                      ReadOnly = false
                    },
                    new DataColumn()
                    {
                      DataType = Type.GetType("System.String"),
                      ColumnName = "tmp_ind",
                      Caption = "Indirizzo",
                      MaxLength = 50,
                      ReadOnly = false
                    },
                    new DataColumn()
                    {
                      DataType = Type.GetType("System.String"),
                      ColumnName = "tmp_loc",
                      Caption = "Località",
                      MaxLength = 50,
                      ReadOnly = false
                    },
                    new DataColumn()
                    {
                      DataType = Type.GetType("System.String"),
                      ColumnName = "tmp_ndo",
                      Caption = "N.documento",
                      MaxLength = 20,
                      ReadOnly = false
                    },
                    new DataColumn()
                    {
                      DataType = Type.GetType("System.DateTime"),
                      ColumnName = "tmp_ddo",
                      Caption = "Data documento",
                      ReadOnly = false
                    },
                    new DataColumn()
                    {
                      DataType = Type.GetType("System.String"),
                      ColumnName = "tmp_tpg",
                      Caption = "Tipo pagamento",
                      MaxLength = 50,
                      ReadOnly = false
                    },
                    new DataColumn()
                    {
                      DataType = Type.GetType("System.Decimal"),
                      ColumnName = "tmp_imp",
                      Caption = "Importo",
                      ReadOnly = false
                    },
                  }
                };
            }
            return t;
        }

        public DataTable TabTmpDocMovDett(string sTab)
        {
            DataTable t = new DataTable(sTab);

            t = new DataTable(sTab)
            {
                Columns = {
                new DataColumn()
                {
                    DataType = Type.GetType("System.String"),
                    ColumnName = "tmp_tri",
                    Caption = "Tipo riga",
                    MaxLength = 2,
                    ReadOnly = false
                },
                new DataColumn()
                {
                    DataType = Type.GetType("System.String"),
                    ColumnName = "tmp_key",
                    Caption = "Key",
                    MaxLength = 20,
                    ReadOnly = false
                },
                new DataColumn()
                {
                    DataType = Type.GetType("System.String"),
                    ColumnName = "tmp_row",
                    Caption = "Riga",
                    MaxLength = 6,
                    ReadOnly = false
                },
                new DataColumn()
                {
                    DataType = Type.GetType("System.DateTime"),
                    ColumnName = "tmp_ddo",
                    Caption = "Data",
                    ReadOnly = false
                },
                new DataColumn()
                {
                    DataType = Type.GetType("System.String"),
                    ColumnName = "tmp_ndo",
                    Caption = "N.documento",
                    MaxLength = 20,
                    ReadOnly = false
                },
                new DataColumn()
                {
                    DataType = Type.GetType("System.String"),
                    ColumnName = "tmp_cod",
                    Caption = "Codice",
                    MaxLength = 7,
                    ReadOnly = false
                },
                new DataColumn()
                {
                    DataType = Type.GetType("System.String"),
                    ColumnName = "tmp_des",
                    Caption = "Descrizione",
                    MaxLength = 100,
                    ReadOnly = false
                },
                new DataColumn()
                {
                    DataType = Type.GetType("System.Decimal"),
                    ColumnName = "tmp_qta",
                    Caption = "Q.ta",
                    ReadOnly = false
                },
                new DataColumn()
                {
                    DataType = Type.GetType("System.Decimal"),
                    ColumnName = "tmp_prz",
                    Caption = "Prezzo",
                    ReadOnly = false
                },
                new DataColumn()
                {
                    DataType = Type.GetType("System.Decimal"),
                    ColumnName = "tmp_imp",
                    Caption = "Importo",
                    ReadOnly = false
                },
                new DataColumn()
                {
                    DataType = Type.GetType("System.String"),
                    ColumnName = "tmp_tmp",
                    Caption = "Tmp",
                    MaxLength = 100,
                    ReadOnly = false
                },
                new DataColumn()
                {
                    DataType = Type.GetType("System.String"),
                    ColumnName = "tmp_for",
                    Caption = "Fornitore",
                    MaxLength = 5,
                    ReadOnly = false
                },
                new DataColumn()
                {
                    DataType = Type.GetType("System.String"),
                    ColumnName = "tmp_fod",
                    Caption = "Fornitore",
                    MaxLength = 70,
                    ReadOnly = false
                },
                new DataColumn()
                {
                    DataType = Type.GetType("System.String"),
                    ColumnName = "tmp_sta",
                    Caption = "Stato",
                    MaxLength = 70,
                    ReadOnly = false
                },
                new DataColumn()
                {
                    DataType = Type.GetType("System.Decimal"),
                    ColumnName = "tmp_cof",
                    Caption = "Costo ultima fattura",
                    ReadOnly = false,
                    DefaultValue = (Decimal)0
                },
                }
            };
            return t;
        }

        public DataTable TabTmpMdfCsv(string sTab)
        {
            DataTable t = new DataTable(sTab);

            if (sTab == "TabTes")
            {
                t = new DataTable(sTab)
                {
                    Columns = {
                    new DataColumn()
                    {
                      DataType = Type.GetType("System.String"),
                      ColumnName = "tab_tab",
                      Caption = "Tabella",
                      MaxLength = 20,
                      ReadOnly = false
                    }
                  }
                };
            }
            else if (sTab == "TabTab")
            {
                t = new DataTable(sTab)
                {
                    Columns = {
                    new DataColumn()
                    {
                      DataType = Type.GetType("System.String"),
                      ColumnName = "tab_tab",
                      Caption = "Nome tabella",
                      MaxLength = 20,
                      ReadOnly = false
                    },
                    new DataColumn()
                    {
                      DataType = Type.GetType("System.String"),
                      ColumnName = "tab_col",
                      Caption = "Nome colonna",
                      MaxLength = 20,
                      ReadOnly = false
                    },
                    new DataColumn()
                    {
                      DataType = Type.GetType("System.String"),
                      ColumnName = "tab_typ",
                      Caption = "Tipo dato",
                      MaxLength = 15,
                      ReadOnly = false
                    },
                    new DataColumn()
                    {
                      DataType = Type.GetType("System.Decimal"),
                      ColumnName = "tab_len",
                      Caption = "Lunghezza campo",
                      ReadOnly = false
                    },
                    new DataColumn()
                    {
                      DataType = Type.GetType("System.Decimal"),
                      ColumnName = "tab_lun",
                      Caption = "Numero interi",
                      ReadOnly = false
                    },
                    new DataColumn()
                    {
                      DataType = Type.GetType("System.Decimal"),
                      ColumnName = "tab_dec",
                      Caption = "Numero decimali",
                      ReadOnly = false
                    },
                    new DataColumn()
                    {
                      DataType = Type.GetType("System.Decimal"),
                      ColumnName = "tab_idn",
                      Caption = "Ident",
                      ReadOnly = false
                    },
                  }
                };
            }
            else if (sTab == "IdxTes")
            {
                t = new DataTable(sTab)
                {
                    Columns = {
                    new DataColumn()
                    {
                      DataType = Type.GetType("System.String"),
                      ColumnName = "idx_tab",
                      Caption = "Nome tabella",
                      MaxLength = 20,
                      ReadOnly = false
                    },
                    new DataColumn()
                    {
                      DataType = Type.GetType("System.String"),
                      ColumnName = "idx_idx",
                      Caption = "Nome indice",
                      MaxLength = 50,
                      ReadOnly = false
                    },
                  }
                };
            }
            else if (sTab == "IdxIdx")
            {
                t = new DataTable(sTab)
                {
                    Columns = {
                    new DataColumn()
                    {
                      DataType = Type.GetType("System.String"),
                      ColumnName = "idx_tab",
                      Caption = "Nome tabella",
                      MaxLength = 20,
                      ReadOnly = false
                    },
                    new DataColumn()
                    {
                      DataType = Type.GetType("System.String"),
                      ColumnName = "idx_idx",
                      Caption = "Nome indice",
                      MaxLength = 50,
                      ReadOnly = false
                    },
                    new DataColumn()
                    {
                      DataType = Type.GetType("System.String"),
                      ColumnName = "idx_col",
                      Caption = "Nome colonna",
                      MaxLength = 50,
                      ReadOnly = false
                    },
                    new DataColumn()
                    {
                      DataType = Type.GetType("System.Decimal"),
                      ColumnName = "idx_xid",
                      Caption = "Posizione",
                      ReadOnly = false
                    },

                  }
                };
            }

            return t;
        }

        public DataTable TabTmpMdfCsvDiff(string sTab)
        {
            DataTable t = new DataTable(sTab);

            t = new DataTable(sTab)
            {
                Columns = {
                    new DataColumn()
                    {
                        DataType = Type.GetType("System.String"),
                        ColumnName = "dif_tip",
                        Caption = "FLD IDX",
                        MaxLength = 3,
                        ReadOnly = false
                    },
                    new DataColumn()
                    {
                        DataType = Type.GetType("System.String"),
                        ColumnName = "dif_tab",
                        Caption = "Nome colonna",
                        MaxLength = 20,
                        ReadOnly = false
                    },
                    new DataColumn()
                    {
                        DataType = Type.GetType("System.String"),
                        ColumnName = "dif_col",
                        Caption = "Nome colonna",
                        MaxLength = 20,
                        ReadOnly = false
                    },
                    new DataColumn()
                    {
                        DataType = Type.GetType("System.String"),
                        ColumnName = "csv_typ",
                        Caption = "tipo data",
                        MaxLength = 15,
                        ReadOnly = false
                    },
                    new DataColumn()
                    {
                        DataType = Type.GetType("System.String"),
                        ColumnName = "mdf_typ",
                        Caption = "tipo data",
                        MaxLength = 15,
                        ReadOnly = false
                    },
                    new DataColumn()
                    {
                        DataType = Type.GetType("System.Decimal"),
                        ColumnName = "csv_len",
                        Caption = "Lunghezza campo",
                        ReadOnly = false
                    },
                    new DataColumn()
                    {
                        DataType = Type.GetType("System.Decimal"),
                        ColumnName = "mdf_len",
                        Caption = "Numero interi",
                        ReadOnly = false
                    },
                    new DataColumn()
                    {
                        DataType = Type.GetType("System.Decimal"),
                        ColumnName = "csv_lun",
                        Caption = "Lunghezza campo",
                        ReadOnly = false
                    },
                    new DataColumn()
                    {
                        DataType = Type.GetType("System.Decimal"),
                        ColumnName = "mdf_lun",
                        Caption = "Numero interi",
                        ReadOnly = false
                    },
                    new DataColumn()
                    {
                        DataType = Type.GetType("System.Decimal"),
                        ColumnName = "csv_dec",
                        Caption = "Numero decimali",
                        ReadOnly = false
                    },
                    new DataColumn()
                    {
                        DataType = Type.GetType("System.Decimal"),
                        ColumnName = "mdf_dec",
                        Caption = "Numero decimali",
                        ReadOnly = false
                    },
                    //new DataColumn()
                    //{
                    //    DataType = Type.GetType("System.String"),
                    //    ColumnName = "csv_fld",
                    //    Caption = "campi",
                    //    MaxLength = 50,
                    //    ReadOnly = false
                    //},
                    //new DataColumn()
                    //{
                    //    DataType = Type.GetType("System.String"),
                    //    ColumnName = "mdf_fld",
                    //    Caption = "campi",
                    //    MaxLength = 50,
                    //    ReadOnly = false
                    //},
                    new DataColumn()
                    {
                        DataType = Type.GetType("System.String"),
                        ColumnName = "csv_idn",
                        Caption = "campi",
                        MaxLength = 50,
                        ReadOnly = false
                    },
                    new DataColumn()
                    {
                        DataType = Type.GetType("System.String"),
                        ColumnName = "mdf_idn",
                        Caption = "campi",
                        MaxLength = 50,
                        ReadOnly = false
                    },
                }
            };

            return t;
        }

        public DataTable TabTmpOrdTermMemor(string sTab)
        {
            DataTable t = new DataTable(sTab);

            t = new DataTable(sTab)
            {
                Columns = {
                    new DataColumn()
                    {
                        DataType = Type.GetType("System.String"),
                        ColumnName = "ord_art",
                        Caption = "Articolo",
                        MaxLength = 7,
                        ReadOnly = false,
                        DefaultValue = (String)""
                    },
                    new DataColumn()
                    {
                        DataType = Type.GetType("System.String"),
                        ColumnName = "ord_ean",
                        Caption = "Ean",
                        MaxLength = 13,
                        ReadOnly = false,
                        DefaultValue = (String)""
                    },
                    new DataColumn()
                    {
                        DataType = Type.GetType("System.String"),
                        ColumnName = "ord_for",
                        Caption = "Fornitore",
                        MaxLength = 5,
                        ReadOnly = false,
                        DefaultValue = (String)""
                    },
                    new DataColumn()
                    {
                        DataType = Type.GetType("System.Decimal"),
                        ColumnName = "ord_qta",
                        Caption = "Qta",
                        ReadOnly = false
                    },
                    new DataColumn()
                    {
                        DataType = Type.GetType("System.String"),
                        ColumnName = "ord_msg",
                        Caption = "Message",
                        MaxLength = 50,
                        ReadOnly = false,
                        DefaultValue = (String)""
                    },
                    new DataColumn()
                    {
                        DataType = Type.GetType("System.String"),
                        ColumnName = "ord_dtt",
                        Caption = "Data ora",
                        MaxLength = 20,
                        ReadOnly = false,
                        DefaultValue = (String)""
                    },
                }
            };

            return t;
        }

        public DataTable TabTmpDivTerm(string sTab)
        {
            DataTable t = new DataTable(sTab);

            t = new DataTable(sTab)
            {
                Columns = {
                    new DataColumn()
                    {
                        DataType = Type.GetType("System.String"),
                        ColumnName = "trm_tip",
                        Caption = "Tipo record T/F",
                        MaxLength = 1,
                        ReadOnly = false,
                        DefaultValue = (String)""
                    },
                    new DataColumn()
                    {
                        DataType = Type.GetType("System.String"),
                        ColumnName = "trm_cod",
                        Caption = "Codice",
                        MaxLength = 6,
                        ReadOnly = false,
                        DefaultValue = (String)""
                    },
                    new DataColumn()
                    {
                        DataType = Type.GetType("System.String"),
                        ColumnName = "trm_div",
                        Caption = "File",
                        MaxLength = 50,
                        ReadOnly = false,
                        DefaultValue = (String)""
                    },
                    new DataColumn()
                    {
                        DataType = Type.GetType("System.String"),
                        ColumnName = "trm_bat",
                        Caption = "File",
                        MaxLength = 50,
                        ReadOnly = false,
                        DefaultValue = (String)""
                    },
                }
            };

            return t;
        }

        public DataTable TabTmpTerRil(string sTab)
        {
            DataTable t = new DataTable(sTab);

            t = new DataTable(sTab)
            {
                Columns = {
                    new DataColumn()
                    {
                        DataType = Type.GetType("System.String"),
                        ColumnName = "tmp_nte",
                        Caption = "Terminalelino",
                        MaxLength = 2,
                        ReadOnly = false,
                        DefaultValue = (String)""
                    },
                    new DataColumn()
                    {
                        DataType = Type.GetType("System.String"),
                        ColumnName = "tmp_nrv",
                        Caption = "Rilevazione",
                        MaxLength = 2,
                        ReadOnly = false,
                        DefaultValue = (String)""
                    },
                }
            };

            return t;
        }

        public DataTable TabTmpStaReparto(string sTab)
        {
            DataTable t = new DataTable(sTab);

            t = new DataTable(sTab)
            {
                Columns = {
                    new DataColumn()
                    {
                        DataType = Type.GetType("System.String"),
                        ColumnName = "tmp_neg",
                        Caption = "Negozio",
                        MaxLength = 3,
                        ReadOnly = false,
                        DefaultValue = (String)""
                    },
                    new DataColumn()
                    {
                        DataType = Type.GetType("System.String"),
                        ColumnName = "tmp_cod",
                        Caption = "Codice",
                        MaxLength = 3,
                        ReadOnly = false,
                        DefaultValue = (String)""
                    },
                    new DataColumn()
                    {
                        DataType = Type.GetType("System.String"),
                        ColumnName = "tmp_des",
                        Caption = "Descrizione",
                        MaxLength = 50,
                        ReadOnly = false,
                        DefaultValue = (String)""
                    },
                    new DataColumn()
                    {
                        DataType = Type.GetType("System.Decimal"),
                        ColumnName = "tmp_imp",
                        Caption = "Importo",
                        ReadOnly = false,
                        DefaultValue = (Decimal)0
                    },
                    new DataColumn()
                    {
                        DataType = Type.GetType("System.Decimal"),
                        ColumnName = "tmp_ils",
                        Caption = "Importo lordo sconto",
                        ReadOnly = false,
                        DefaultValue = (Decimal)0
                    },
                    new DataColumn()
                    {
                        DataType = Type.GetType("System.Decimal"),
                        ColumnName = "tmp_inc",
                        Caption = "Incidenza",
                        ReadOnly = false,
                        DefaultValue = (Decimal)0
                    },
                    new DataColumn()
                    {
                        DataType = Type.GetType("System.Decimal"),
                        ColumnName = "tmp_sco",
                        Caption = "Scontrini",
                        ReadOnly = false,
                        DefaultValue = (Decimal)0
                    },
                    new DataColumn()
                    {
                        DataType = Type.GetType("System.Decimal"),
                        ColumnName = "tmp_sci",
                        Caption = "Incidenza",
                        ReadOnly = false,
                        DefaultValue = (Decimal)0
                    },
                    new DataColumn()
                    {
                        DataType = Type.GetType("System.Decimal"),
                        ColumnName = "tmp_scm",
                        Caption = "Val.medio",
                        ReadOnly = false,
                        DefaultValue = (Decimal)0
                    },
                    new DataColumn()
                    {
                        DataType = Type.GetType("System.Decimal"),
                        ColumnName = "tmp_qta",
                        Caption = "Pezzi",
                        ReadOnly = false,
                        DefaultValue = (Decimal)0
                    },
                    new DataColumn()
                    {
                        DataType = Type.GetType("System.Decimal"),
                        ColumnName = "tmp_qti",
                        Caption = "Incidenza",
                        ReadOnly = false,
                        DefaultValue = (Decimal)0
                    },
                    new DataColumn()
                    {
                        DataType = Type.GetType("System.Decimal"),
                        ColumnName = "tmp_qtm",
                        Caption = "Val.medio",
                        ReadOnly = false,
                        DefaultValue = (Decimal)0
                    },
                    new DataColumn()
                    {
                        DataType = Type.GetType("System.Decimal"),
                        ColumnName = "tmp_niv",
                        Caption = "Netto IVA",
                        ReadOnly = false,
                        DefaultValue = (Decimal)0
                    },
                    new DataColumn()
                    {
                        DataType = Type.GetType("System.Decimal"),
                        ColumnName = "tmp_vcc",
                        Caption = "Venduto con costo",
                        ReadOnly = false,
                        DefaultValue = (Decimal)0
                    },
                    new DataColumn()
                    {
                        DataType = Type.GetType("System.Decimal"),
                        ColumnName = "tmp_cdv",
                        Caption = "Costo del venduto",
                        ReadOnly = false,
                        DefaultValue = (Decimal)0
                    },
                    new DataColumn()
                    {
                        DataType = Type.GetType("System.Decimal"),
                        ColumnName = "tmp_mav",
                        Caption = "Margine venduto",
                        ReadOnly = false,
                        DefaultValue = (Decimal)0
                    },
                }
            };

            return t;
        }

        public DataTable TabTmpStaEcr(string sTab)
        {
            DataTable t = new DataTable(sTab);

            t = new DataTable(sTab)
            {
                Columns = {
                    new DataColumn()
                    {
                        DataType = Type.GetType("System.String"),
                        ColumnName = "tmp_neg",
                        Caption = "Negozio",
                        MaxLength = 3,
                        ReadOnly = false,
                        DefaultValue = (String)""
                    },
                    new DataColumn()
                    {
                        DataType = Type.GetType("System.String"),
                        ColumnName = "tmp_cod",
                        Caption = "lv1+lv2+lv3",
                        MaxLength = 9,
                        ReadOnly = false,
                        DefaultValue = (String)""
                    },
                    new DataColumn()
                    {
                        DataType = Type.GetType("System.String"),
                        ColumnName = "tmp_ec1",
                        Caption = "Codice",
                        MaxLength = 3,
                        ReadOnly = false,
                        DefaultValue = (String)""
                    },
                    new DataColumn()
                    {
                        DataType = Type.GetType("System.String"),
                        ColumnName = "tmp_ed1",
                        Caption = "Codice",
                        MaxLength = 30,
                        ReadOnly = false,
                        DefaultValue = (String)""
                    },
                    new DataColumn()
                    {
                        DataType = Type.GetType("System.String"),
                        ColumnName = "tmp_ec2",
                        Caption = "Codice",
                        MaxLength = 3,
                        ReadOnly = false,
                        DefaultValue = (String)""
                    },
                    new DataColumn()
                    {
                        DataType = Type.GetType("System.String"),
                        ColumnName = "tmp_ed2",
                        Caption = "Codice",
                        MaxLength = 30,
                        ReadOnly = false,
                        DefaultValue = (String)""
                    },
                    new DataColumn()
                    {
                        DataType = Type.GetType("System.String"),
                        ColumnName = "tmp_ec3",
                        Caption = "Codice",
                        MaxLength = 3,
                        ReadOnly = false,
                        DefaultValue = (String)""
                    },
                    new DataColumn()
                    {
                        DataType = Type.GetType("System.String"),
                        ColumnName = "tmp_ed3",
                        Caption = "Codice",
                        MaxLength = 30,
                        ReadOnly = false,
                        DefaultValue = (String)""
                    },
                    new DataColumn()
                    {
                        DataType = Type.GetType("System.String"),
                        ColumnName = "tmp_des",
                        Caption = "Descrizione",
                        MaxLength = 50,
                        ReadOnly = false,
                        DefaultValue = (String)""
                    },
                    new DataColumn()
                    {
                        DataType = Type.GetType("System.Decimal"),
                        ColumnName = "tmp_imp",
                        Caption = "Importo",
                        ReadOnly = false,
                        DefaultValue = (Decimal)0
                    },
                    new DataColumn()
                    {
                        DataType = Type.GetType("System.Decimal"),
                        ColumnName = "tmp_ils",
                        Caption = "Importo lordo sconto",
                        ReadOnly = false,
                        DefaultValue = (Decimal)0
                    },
                    new DataColumn()
                    {
                        DataType = Type.GetType("System.Decimal"),
                        ColumnName = "tmp_inc",
                        Caption = "Incidenza",
                        ReadOnly = false,
                        DefaultValue = (Decimal)0
                    },
                    new DataColumn()
                    {
                        DataType = Type.GetType("System.Decimal"),
                        ColumnName = "tmp_sco",
                        Caption = "Scontrini",
                        ReadOnly = false,
                        DefaultValue = (Decimal)0
                    },
                    new DataColumn()
                    {
                        DataType = Type.GetType("System.Decimal"),
                        ColumnName = "tmp_sci",
                        Caption = "Incidenza",
                        ReadOnly = false,
                        DefaultValue = (Decimal)0
                    },
                    new DataColumn()
                    {
                        DataType = Type.GetType("System.Decimal"),
                        ColumnName = "tmp_scm",
                        Caption = "Val.medio",
                        ReadOnly = false,
                        DefaultValue = (Decimal)0
                    },
                    new DataColumn()
                    {
                        DataType = Type.GetType("System.Decimal"),
                        ColumnName = "tmp_qta",
                        Caption = "Pezzi",
                        ReadOnly = false,
                        DefaultValue = (Decimal)0
                    },
                    new DataColumn()
                    {
                        DataType = Type.GetType("System.Decimal"),
                        ColumnName = "tmp_qti",
                        Caption = "Incidenza",
                        ReadOnly = false,
                        DefaultValue = (Decimal)0
                    },
                    new DataColumn()
                    {
                        DataType = Type.GetType("System.Decimal"),
                        ColumnName = "tmp_qtm",
                        Caption = "Val.medio",
                        ReadOnly = false,
                        DefaultValue = (Decimal)0
                    },
                    new DataColumn()
                    {
                        DataType = Type.GetType("System.Decimal"),
                        ColumnName = "tmp_niv",
                        Caption = "Netto IVA",
                        ReadOnly = false,
                        DefaultValue = (Decimal)0
                    },
                    new DataColumn()
                    {
                        DataType = Type.GetType("System.Decimal"),
                        ColumnName = "tmp_vcc",
                        Caption = "Venduto con costo",
                        ReadOnly = false,
                        DefaultValue = (Decimal)0
                    },
                    new DataColumn()
                    {
                        DataType = Type.GetType("System.Decimal"),
                        ColumnName = "tmp_cdv",
                        Caption = "Costo del venduto",
                        ReadOnly = false,
                        DefaultValue = (Decimal)0
                    },
                    new DataColumn()
                    {
                        DataType = Type.GetType("System.Decimal"),
                        ColumnName = "tmp_mav",
                        Caption = "Margine venduto",
                        ReadOnly = false,
                        DefaultValue = (Decimal)0
                    },
                }
            };

            return t;
        }

        public DataTable TabTmpArtPluBilance(string sTab)
        {
            DataTable t = new DataTable(sTab);

            t = new DataTable(sTab)
            {
                Columns = {
                    new DataColumn()
                    {
                        DataType = Type.GetType("System.String"),
                        ColumnName = "tmp_art",
                        Caption = "Codice",
                        MaxLength = 7,
                        ReadOnly = false,
                        DefaultValue = (String)""
                    },
                    new DataColumn()
                    {
                        DataType = Type.GetType("System.String"),
                        ColumnName = "tmp_des",
                        Caption = "Descrizione",
                        MaxLength = 50,
                        ReadOnly = false,
                        DefaultValue = (String)""
                    },
                    new DataColumn()
                    {
                        DataType = Type.GetType("System.String"),
                        ColumnName = "tmp_plu",
                        Caption = "PLU",
                        MaxLength = 6,
                        ReadOnly = false,
                        DefaultValue = (String)""
                    },
                    new DataColumn()
                    {
                        DataType = Type.GetType("System.String"),
                        ColumnName = "tmp_tas",
                        Caption = "Tasto bilancia",
                        MaxLength = 4,
                        ReadOnly = false,
                        DefaultValue = (String)""
                    },
                    new DataColumn()
                    {
                        DataType = Type.GetType("System.Decimal"),
                        ColumnName = "tmp_cos",
                        Caption = "Costo",
                        //MaxLength = 50,
                        ReadOnly = false,
                        DefaultValue = (Decimal)0
                    },
                    new DataColumn()
                    {
                        DataType = Type.GetType("System.Decimal"),
                        ColumnName = "tmp_pve",
                        Caption = "Prezzo",
                        //MaxLength = 50,
                        ReadOnly = false,
                        DefaultValue = (Decimal)0
                    },
                    new DataColumn()
                    {
                        DataType = Type.GetType("System.String"),
                        ColumnName = "tmp_sta",
                        Caption = "Stato",
                        MaxLength = 1,
                        ReadOnly = false,
                        DefaultValue = (String)""
                    },
                    new DataColumn()
                    {
                        DataType = Type.GetType("System.String"),
                        ColumnName = "tmp_ean",
                        Caption = "Barcode",
                        MaxLength = 13,
                        ReadOnly = false,
                        DefaultValue = (String)""
                    },
                    new DataColumn()
                    {
                        DataType = Type.GetType("System.String"),
                        ColumnName = "tmp_rep",
                        Caption = "Reparto",
                        MaxLength = 3,
                        ReadOnly = false,
                        DefaultValue = (String)""
                    },
                    new DataColumn()
                    {
                        DataType = Type.GetType("System.String"),
                        ColumnName = "tmp_red",
                        Caption = "Reparto",
                        MaxLength = 20,
                        ReadOnly = false,
                        DefaultValue = (String)""
                    },
                }
            };

            return t;
        }

        public DataTable TabTmpVenErrori(string sTab)
        {
            DataTable t = new DataTable(sTab);

            t = new DataTable(sTab)
            {
                Columns = {
                    new DataColumn()
                    {
                        DataType = Type.GetType("System.String"),
                        ColumnName = "tmp_ean",
                        Caption = "Codice",
                        MaxLength = 13,
                        ReadOnly = false,
                        DefaultValue = (String)""
                    },
                    new DataColumn()
                    {
                        DataType = Type.GetType("System.String"),
                        ColumnName = "tmp_msg",
                        Caption = "Messaggio",
                        MaxLength = 100,
                        ReadOnly = false,
                        DefaultValue = (String)""
                    },
                    //new DataColumn()
                    //{
                    //    DataType = Type.GetType("System.Decimal"),
                    //    ColumnName = "tmp_imp",
                    //    Caption = "Descrizione",
                    //    //MaxLength = 50,
                    //    ReadOnly = false,
                    //    DefaultValue = (Decimal)0
                    //},
                }
            };

            return t;
        }

        public DataTable TabTmpArtVenduto(string sTab)
        {
            DataTable t = new DataTable(sTab);

            t = new DataTable(sTab)
            {
                Columns = {
                    new DataColumn()
                    {
                        DataType = Type.GetType("System.String"),
                        ColumnName = "tmp_art",
                        Caption = "Codice",
                        MaxLength = 20,
                        ReadOnly = false,
                        DefaultValue = (String)""
                    },
                    new DataColumn()
                    {
                        DataType = Type.GetType("System.String"),
                        ColumnName = "tmp_ard",
                        Caption = "Descrizione",
                        MaxLength = 50,
                        ReadOnly = false,
                        DefaultValue = (String)""
                    },
                    new DataColumn()
                    {
                        DataType = Type.GetType("System.String"),
                        ColumnName = "tmp_rep",
                        Caption = "Reparto",
                        MaxLength = 3,
                        ReadOnly = false,
                        DefaultValue = (String)""
                    },
                    new DataColumn()
                    {
                        DataType = Type.GetType("System.String"),
                        ColumnName = "tmp_red",
                        Caption = "Reparto des",
                        MaxLength = 50,
                        ReadOnly = false,
                        DefaultValue = (String)""
                    },
                    new DataColumn()
                    {
                        DataType = Type.GetType("System.String"),
                        ColumnName = "tmp_sgn",
                        Caption = "Segno",
                        MaxLength = 50,
                        ReadOnly = false,
                        DefaultValue = (String)""
                    },
                    new DataColumn()
                    {
                        DataType = Type.GetType("System.Decimal"),
                        ColumnName = "tmp_qta",
                        Caption = "Q.tà",
                        ReadOnly = false,
                        DefaultValue = (Decimal)0
                    },
                    new DataColumn()
                    {
                        DataType = Type.GetType("System.Decimal"),
                        ColumnName = "tmp_qve",
                        Caption = "Q.tà",
                        ReadOnly = false,
                        DefaultValue = (Decimal)0
                    },
                    new DataColumn()
                    {
                        DataType = Type.GetType("System.Decimal"),
                        ColumnName = "tmp_qfa",
                        Caption = "Q.tà fatture",
                        ReadOnly = false,
                        DefaultValue = (Decimal)0
                    },
                    new DataColumn()
                    {
                        DataType = Type.GetType("System.Decimal"),
                        ColumnName = "tmp_qmo",
                        Caption = "Q.tà movimenti",
                        ReadOnly = false,
                        DefaultValue = (Decimal)0
                    },
                    new DataColumn()
                    {
                        DataType = Type.GetType("System.Decimal"),
                        ColumnName = "tmp_pco",
                        Caption = "Prezzo costo",
                        ReadOnly = false,
                        DefaultValue = (Decimal)0
                    },
                    new DataColumn()
                    {
                        DataType = Type.GetType("System.Decimal"),
                        ColumnName = "tmp_pve",
                        Caption = "Prezzo vendita",
                        ReadOnly = false,
                        DefaultValue = (Decimal)0
                    },
                    new DataColumn()
                    {
                        DataType = Type.GetType("System.Decimal"),
                        ColumnName = "tmp_vve",
                        Caption = "Valore venduto",
                        ReadOnly = false,
                        DefaultValue = (Decimal)0
                    },
                    new DataColumn()
                    {
                        DataType = Type.GetType("System.Decimal"),
                        ColumnName = "tmp_val",
                        Caption = "Valore totale",
                        ReadOnly = false,
                        DefaultValue = (Decimal)0
                    },
                    new DataColumn()
                    {
                        DataType = Type.GetType("System.Decimal"),
                        ColumnName = "tmp_vfa",
                        Caption = "Valore acquistato",
                        ReadOnly = false,
                        DefaultValue = (Decimal)0
                    },
                    new DataColumn()
                    {
                        DataType = Type.GetType("System.Decimal"),
                        ColumnName = "tmp_vmo",
                        Caption = "Valore movimentato",
                        ReadOnly = false,
                        DefaultValue = (Decimal)0
                    },
                    new DataColumn()
                    {
                        DataType = Type.GetType("System.String"),
                        ColumnName = "tmp_ean",
                        Caption = "Ean",
                        MaxLength = 13,
                        ReadOnly = false,
                        DefaultValue = (String)""
                    },

                }
            };

            return t;
        }

        public DataTable TabTmpArtCtrl(string sTab)
        {
            DataTable t = new DataTable(sTab);

            t = new DataTable(sTab)
            {
                Columns = {
                    new DataColumn()
                    {
                        DataType = Type.GetType("System.String"),
                        ColumnName = "tmp_ean",
                        Caption = "Ean",
                        MaxLength = 13,
                        ReadOnly = false,
                        DefaultValue = (String)""
                    },
                    new DataColumn()
                    {
                        DataType = Type.GetType("System.String"),
                        ColumnName = "tmp_ar1",
                        Caption = "Codice",
                        MaxLength = 20,
                        ReadOnly = false,
                        DefaultValue = (String)""
                    },
                    new DataColumn()
                    {
                        DataType = Type.GetType("System.String"),
                        ColumnName = "tmp_de1",
                        Caption = "Descrizione",
                        MaxLength = 50,
                        ReadOnly = false,
                        DefaultValue = (String)""
                    },
                    new DataColumn()
                    {
                        DataType = Type.GetType("System.String"),
                        ColumnName = "tmp_ar2",
                        Caption = "Codice",
                        MaxLength = 20,
                        ReadOnly = false,
                        DefaultValue = (String)""
                    },
                    new DataColumn()
                    {
                        DataType = Type.GetType("System.String"),
                        ColumnName = "tmp_de2",
                        Caption = "Descrizione",
                        MaxLength = 50,
                        ReadOnly = false,
                        DefaultValue = (String)""
                    },
                    new DataColumn()
                    {
                        DataType = Type.GetType("System.Boolean"),
                        ColumnName = "tmp_ann",
                        Caption = "Annullato",
                        //MaxLength = 50,
                        ReadOnly = false,
                        DefaultValue = (Boolean)false
                    },
                }
            };

            return t;
        }

        public DataTable TabTmpArtCtrl2(string sTab)
        {
            DataTable t = new DataTable(sTab);

            t = new DataTable(sTab)
            {
                Columns = {
                    new DataColumn()
                    {
                        DataType = Type.GetType("System.String"),
                        ColumnName = "tmp_art",
                        Caption = "Codice",
                        MaxLength = 10,
                        ReadOnly = false,
                        DefaultValue = (String)""
                    },
                    new DataColumn()
                    {
                        DataType = Type.GetType("System.String"),
                        ColumnName = "tmp_des",
                        Caption = "Descrizione",
                        MaxLength = 50,
                        ReadOnly = false,
                        DefaultValue = (String)""
                    },
                    new DataColumn()
                    {
                        DataType = Type.GetType("System.String"),
                        ColumnName = "tmp_sta",
                        Caption = "Stato",
                        MaxLength = 1,
                        ReadOnly = false,
                        DefaultValue = (String)""
                    },
                    new DataColumn()
                    {
                        DataType = Type.GetType("System.String"),
                        ColumnName = "mdb_art",
                        Caption = "Codice",
                        MaxLength = 15,
                        ReadOnly = false,
                        DefaultValue = (String)""
                    },
                    new DataColumn()
                    {
                        DataType = Type.GetType("System.String"),
                        ColumnName = "mdb_des",
                        Caption = "Descrizione",
                        MaxLength = 50,
                        ReadOnly = false,
                        DefaultValue = (String)""
                    },

                }
            };

            return t;
        }

        public DataTable TabTmpVenCtrl(string sTab)
        {
            DataTable t = new DataTable(sTab);

            t = new DataTable(sTab)
            {
                Columns = {
                    new DataColumn()
                    {
                        DataType = Type.GetType("System.String"),
                        ColumnName = "tmp_cau",
                        Caption = "Cau",
                        MaxLength = 3,
                        ReadOnly = false,
                        DefaultValue = (String)""
                    },
                    new DataColumn()
                    {
                        DataType = Type.GetType("System.String"),
                        ColumnName = "tmp_pos",
                        Caption = "Pos",
                        MaxLength = 2,
                        ReadOnly = false,
                        DefaultValue = (String)""
                    },
                    new DataColumn()
                    {
                        DataType = Type.GetType("System.String"),
                        ColumnName = "tmp_sco",
                        Caption = "Scontrino",
                        MaxLength = 5,
                        ReadOnly = false,
                        DefaultValue = (String)""
                    },
                    new DataColumn()
                    {
                        DataType = Type.GetType("System.Decimal"),
                        ColumnName = "tmp_vet",
                        Caption = "Q.tà",
                        ReadOnly = false,
                        DefaultValue = (Decimal)0
                    },
                    new DataColumn()
                    {
                        DataType = Type.GetType("System.Decimal"),
                        ColumnName = "tmp_ven",
                        Caption = "venduto x dettaglio",
                        ReadOnly = false,
                        DefaultValue = (Decimal)0
                    },
                    new DataColumn()
                    {
                        DataType = Type.GetType("System.Decimal"),
                        ColumnName = "tmp_vep",
                        Caption = "Pagamenti",
                        ReadOnly = false,
                        DefaultValue = (Decimal)0
                    },
                    new DataColumn()
                    {
                        DataType = Type.GetType("System.Decimal"),
                        ColumnName = "tmp_vnd",
                        Caption = "Diff",
                        ReadOnly = false,
                        DefaultValue = (Decimal)0
                    },
                }
            };

            return t;
        }

        public DataTable TabTmpArtGiacenza(string sTab)
        {
            DataTable t = new DataTable(sTab);

            t = new DataTable(sTab)
            {
                Columns = {
                    new DataColumn()
                    {
                        DataType = Type.GetType("System.String"),
                        ColumnName = "tmp_art",
                        Caption = "Art",
                        MaxLength = 7,
                        ReadOnly = false,
                        DefaultValue = (String)""
                    },
                    new DataColumn()
                    {
                        DataType = Type.GetType("System.String"),
                        ColumnName = "tmp_ard",
                        Caption = "Descrizione",
                        MaxLength = 50,
                        ReadOnly = false,
                        DefaultValue = (String)""
                    },
                    new DataColumn()
                    {
                        DataType = Type.GetType("System.String"),
                        ColumnName = "tmp_rep",
                        Caption = "Reparto",
                        MaxLength = 3,
                        ReadOnly = false,
                        DefaultValue = (String)""
                    },
                    new DataColumn()
                    {
                        DataType = Type.GetType("System.String"),
                        ColumnName = "tmp_red",
                        Caption = "Reparto",
                        MaxLength = 30,
                        ReadOnly = false,
                        DefaultValue = (String)""
                    },
                    new DataColumn()
                    {
                        DataType = Type.GetType("System.Decimal"),
                        ColumnName = "tmp_qac",
                        Caption = "Q.tà",
                        ReadOnly = false,
                        DefaultValue = (Decimal)0
                    },
                    new DataColumn()
                    {
                        DataType = Type.GetType("System.Decimal"),
                        ColumnName = "tmp_qve",
                        Caption = "venduto",
                        ReadOnly = false,
                        DefaultValue = (Decimal)0
                    },
                    new DataColumn()
                    {
                        DataType = Type.GetType("System.Decimal"),
                        ColumnName = "tmp_qmo",
                        Caption = "Mov. rettifica",
                        ReadOnly = false,
                        DefaultValue = (Decimal)0
                    },
                    new DataColumn()
                    {
                        DataType = Type.GetType("System.Decimal"),
                        ColumnName = "tmp_qta",
                        Caption = "Giacenza",
                        ReadOnly = false,
                        DefaultValue = (Decimal)0
                    },
                    new DataColumn()
                    {
                        DataType = Type.GetType("System.Decimal"),
                        ColumnName = "tmp_cos",
                        Caption = "Costo",
                        ReadOnly = false,
                        DefaultValue = (Decimal)0
                    },
                    new DataColumn()
                    {
                        DataType = Type.GetType("System.Decimal"),
                        ColumnName = "tmp_val",
                        Caption = "Valore",
                        ReadOnly = false,
                        DefaultValue = (Decimal)0
                    },
                }
            };

            return t;
        }

        public DataTable TabTmpStatistiche(string sTab)
        {
            DataTable t = new DataTable(sTab);

            t = new DataTable(sTab)
            {
                Columns = {
                    new DataColumn()
                    {
                        DataType = Type.GetType("System.String"),
                        ColumnName = "tmp_odx",
                        Caption = "Ordinamento",
                        MaxLength = 1,
                        ReadOnly = false,
                        DefaultValue = (String)""
                    },
                    new DataColumn()
                    {
                        DataType = Type.GetType("System.String"),
                        ColumnName = "tmp_key",
                        Caption = "Data",
                        ReadOnly = false,
                    },
                    new DataColumn()
                    {
                        DataType = Type.GetType("System.DateTime"),
                        ColumnName = "tmp_day",
                        Caption = "Data",
                        ReadOnly = false,
                    },
                    new DataColumn()
                    {
                        DataType = Type.GetType("System.String"),
                        ColumnName = "tmp_mag",
                        Caption = "Magazzino",
                        MaxLength = 2,
                        ReadOnly = false,
                        DefaultValue = (String)""
                    },
                    new DataColumn()
                    {
                        DataType = Type.GetType("System.String"),
                        ColumnName = "tmp_neg",
                        Caption = "Negozio",
                        MaxLength = 15,
                        ReadOnly = false,
                        DefaultValue = (String)""
                    },
                    new DataColumn()
                    {
                        DataType = Type.GetType("System.String"),
                        ColumnName = "tmp_ndo",
                        Caption = "Documento",
                        MaxLength = 15,
                        ReadOnly = false,
                        DefaultValue = (String)""
                    },
                    new DataColumn()
                    {
                        DataType = Type.GetType("System.String"),
                        ColumnName = "tmp_tre",
                        Caption = "Tipo record",
                        MaxLength = 3,
                        ReadOnly = false,
                        DefaultValue = (String)""
                    },
                    new DataColumn()
                    {
                        DataType = Type.GetType("System.String"),
                        ColumnName = "tmp_des",
                        Caption = "Descrizione",
                        MaxLength = 150,
                        ReadOnly = false,
                        DefaultValue = (String)""
                    },
                    new DataColumn()
                    {
                        DataType = Type.GetType("System.Decimal"),
                        ColumnName = "tmp_imp",
                        Caption = "Importo",
                        ReadOnly = false,
                        DefaultValue = (Decimal)0
                    },
                }
            };

            return t;
        }

        public DataTable TabTmpStaArt(string sTab)
        {
            DataTable t = new DataTable(sTab);
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "tmp_neg",
                Caption = "negozio",
                MaxLength = 3,
                ReadOnly = false,
                DefaultValue = (String)""
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "tmp_tip",
                Caption = "Tipo",
                MaxLength = 3,
                ReadOnly = false,
                DefaultValue = (String)""
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "tmp_art",
                Caption = "Art",
                MaxLength = 7,
                ReadOnly = false,
                DefaultValue = (String)""
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "tmp_ard",
                Caption = "Descrizione",
                MaxLength = 50,
                ReadOnly = false,
                DefaultValue = (String)""
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.DateTime"),
                ColumnName = "tmp_day",
                Caption = "Data",
                ReadOnly = false,
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.Decimal"),
                ColumnName = "tmp_qta",
                Caption = "Q.tà",
                ReadOnly = false,
                DefaultValue = (Decimal)0
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.Decimal"),
                ColumnName = "tmp_qkg",
                Caption = "Q.tà peso",
                ReadOnly = false,
                DefaultValue = (Decimal)0
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.Decimal"),
                ColumnName = "tmp_cos",
                Caption = "Acquistato",
                ReadOnly = false,
                DefaultValue = (Decimal)0
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.Decimal"),
                ColumnName = "tmp_acq",
                Caption = "Acquistato",
                ReadOnly = false,
                DefaultValue = (Decimal)0
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.Decimal"),
                ColumnName = "tmp_prv",
                Caption = "Venduto",
                ReadOnly = false,
                DefaultValue = (Decimal)0
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.Decimal"),
                ColumnName = "tmp_ven",
                Caption = "Venduto",
                ReadOnly = false,
                DefaultValue = (Decimal)0
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.Decimal"),
                ColumnName = "tmp_mov",
                Caption = "Movimenti",
                ReadOnly = false,
                DefaultValue = (Decimal)0
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "tmp_des",
                Caption = "Descrizione",
                MaxLength = 150,
                ReadOnly = false,
                DefaultValue = (String)""
            });

            return t;
        }

        public DataTable TabTmpAnaArtPVendita(string sTab)
        {
            DataTable t = new DataTable(sTab);
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "LivIdx",
                Caption = "Idx",
                MaxLength = 3,
                ReadOnly = false,
                DefaultValue = (String)""
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "LivLis",
                Caption = "Listino",
                MaxLength = 3,
                ReadOnly = false,
                DefaultValue = (String)""
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "LivTip",
                Caption = "Tipo",
                MaxLength = 2,
                ReadOnly = false,
                DefaultValue = (String)""
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.Decimal"),
                ColumnName = "LivPrv",
                Caption = "Prezzo",
                ReadOnly = false,
                DefaultValue = (Decimal)0
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.Decimal"),
                ColumnName = "LivMav",
                Caption = "Margine valore",
                ReadOnly = false,
                DefaultValue = (Decimal)0
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.Decimal"),
                ColumnName = "LivMap",
                Caption = "Margine %",
                ReadOnly = false,
                DefaultValue = (Decimal)0
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.Decimal"),
                ColumnName = "LivMao",
                Caption = "Margine % obbiettivo",
                ReadOnly = false,
                DefaultValue = (Decimal)0
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.DateTime"),
                ColumnName = "LivDti",
                Caption = "Data inizio",
                ReadOnly = false
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.DateTime"),
                ColumnName = "LivDtf",
                Caption = "Data inizio",
                ReadOnly = false
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.Boolean"),
                ColumnName = "LivAnn",
                Caption = "Annullato",
                ReadOnly = false,
                DefaultValue = (Boolean)false
            });

            return t;
        }

        public DataTable TabTmpDivArtDif(string sTab)
        {
            DataTable t = new DataTable(sTab);
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "tmp_div",
                Caption = "Articolo",
                MaxLength = 3,
                ReadOnly = false,
                DefaultValue = (String)""
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "tmp_art",
                Caption = "Articolo",
                MaxLength = 7,
                ReadOnly = false,
                DefaultValue = (String)""
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "tmp_tip",
                Caption = "Ana o var",
                MaxLength = 3,
                ReadOnly = false,
                DefaultValue = (String)""
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "tmp_ard",
                Caption = "Descrizione",
                MaxLength = 50,
                ReadOnly = false,
                DefaultValue = (String)""
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "tmp_umi",
                Caption = "UMI",
                MaxLength = 2,
                ReadOnly = false,
                DefaultValue = (String)""
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.Decimal"),
                ColumnName = "tmp_pne",
                Caption = "Peso",
                ReadOnly = false,
                DefaultValue = (Decimal)0
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.Decimal"),
                ColumnName = "tmp_pxc",
                Caption = "Pxc",
                ReadOnly = false,
                DefaultValue = (Decimal)0
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "tmp_plu",
                Caption = "PLU",
                MaxLength = 7,
                ReadOnly = false,
                DefaultValue = (String)""
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "tmp_rep",
                Caption = "Reparto",
                MaxLength = 3,
                ReadOnly = false,
                DefaultValue = (String)""
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "tmp_reb",
                Caption = "Reparto bilancia",
                MaxLength = 1,
                ReadOnly = false,
                DefaultValue = (String)""
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "tmp_ori",
                Caption = "Origine",
                MaxLength = 7,
                ReadOnly = false,
                DefaultValue = (String)""
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "tmp_cat",
                Caption = "Categoria",
                MaxLength = 3,
                ReadOnly = false,
                DefaultValue = (String)""
            });

            return t;

        }

        public DataTable TabTmpDivArtPrn(string sTab)
        {
            DataTable t = new DataTable(sTab);
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "tmp_art",
                Caption = "Articolo",
                MaxLength = 7,
                ReadOnly = false,
                DefaultValue = (String)""
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "tmp_arf",
                Caption = "Articolo fornitore",
                MaxLength = 15,
                ReadOnly = false,
                DefaultValue = (String)""
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "tmp_ard",
                Caption = "Descrizione",
                MaxLength = 50,
                ReadOnly = false,
                DefaultValue = (String)""
            });
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
                ColumnName = "tmp_prp",
                Caption = "Prezzo",
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
                ColumnName = "tmp_plu",
                Caption = "PLU",
                MaxLength = 6,
                ReadOnly = false,
                DefaultValue = (String)""
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "tmp_sta",
                Caption = "Stato",
                MaxLength = 10,
                ReadOnly = false,
                DefaultValue = (String)""
            });

            return t;

        }

        public DataTable TabTmpJounal(string sTab)
        {
            DataTable t = new DataTable(sTab);
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "tmp_day",
                Caption = "Data",
                MaxLength = 10,
                ReadOnly = false,
                DefaultValue = (String)""
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "tmp_ora",
                Caption = "Data",
                MaxLength = 5,
                ReadOnly = false,
                DefaultValue = (String)""
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "tmp_sco",
                Caption = "Scontrino",
                MaxLength = 5,
                ReadOnly = false,
                DefaultValue = (String)""
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "tmp_num",
                Caption = "Numero progressivo",
                MaxLength = 5,
                ReadOnly = false,
                DefaultValue = (String)""
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.Decimal"),
                ColumnName = "tmp_imj",
                Caption = "Importo journal",
                ReadOnly = false,
                DefaultValue = (Decimal)0
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.Decimal"),
                ColumnName = "tmp_ims",
                Caption = "Importo statistiche",
                ReadOnly = false,
                DefaultValue = (Decimal)0
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.Decimal"),
                ColumnName = "tmp_dif",
                Caption = "Differenza",
                ReadOnly = false,
                DefaultValue = (Decimal)0
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.Decimal"),
                ColumnName = "tmp_i_A",
                Caption = "IVA A",
                ReadOnly = false,
                DefaultValue = (Decimal)0
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.Decimal"),
                ColumnName = "tmp_i_B",
                Caption = "IVA B",
                ReadOnly = false,
                DefaultValue = (Decimal)0
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.Decimal"),
                ColumnName = "tmp_i_C",
                Caption = "IVA C",
                ReadOnly = false,
                DefaultValue = (Decimal)0
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.Decimal"),
                ColumnName = "tmp_i_D",
                Caption = "IVA D",
                ReadOnly = false,
                DefaultValue = (Decimal)0
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.Decimal"),
                ColumnName = "tmp_s_A",
                Caption = "IVA A",
                ReadOnly = false,
                DefaultValue = (Decimal)0
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.Decimal"),
                ColumnName = "tmp_s_B",
                Caption = "IVA B",
                ReadOnly = false,
                DefaultValue = (Decimal)0
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.Decimal"),
                ColumnName = "tmp_s_C",
                Caption = "IVA C",
                ReadOnly = false,
                DefaultValue = (Decimal)0
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.Decimal"),
                ColumnName = "tmp_s_D",
                Caption = "IVA D",
                ReadOnly = false,
                DefaultValue = (Decimal)0
            });

            return t;
        }

        public DataTable TabTmpRows(string sTab)
        {
            DataTable t = new DataTable(sTab);
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "tmp_art",
                Caption = "Codice articolo",
                MaxLength = 7,
                ReadOnly = false,
                DefaultValue = (String)""
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "tmp_umi",
                Caption = "UM",
                MaxLength = 2,
                ReadOnly = false,
                DefaultValue = (String)""
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "tmp_pxc",
                Caption = "PxC",
                MaxLength = 5,
                ReadOnly = false,
                DefaultValue = (String)""
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "tmp_lot",
                Caption = "Lotto",
                MaxLength = 15,
                ReadOnly = false,
                DefaultValue = (String)""
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.Decimal"),
                ColumnName = "tmp_tar",
                Caption = "Q.tà peso",
                ReadOnly = false,
                DefaultValue = (Decimal)0
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.Decimal"),
                ColumnName = "tmp_qkg",
                Caption = "Q.tà peso",
                ReadOnly = false,
                DefaultValue = (Decimal)0
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.Decimal"),
                ColumnName = "tmp_pve",
                Caption = "Prezzo",
                ReadOnly = false,
                DefaultValue = (Decimal)0
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.Decimal"),
                ColumnName = "tmp_prp",
                Caption = "Prezzo pubblico",
                ReadOnly = false,
                DefaultValue = (Decimal)0
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.Decimal"),
                ColumnName = "tmp_gsc",
                Caption = "Giorni scadenza",
                ReadOnly = false,
                DefaultValue = (Decimal)0
            });

            return t;
        }

        public DataTable TabTmpArtRotazione(string sTab, int intIni, int intFin)
        {
            DataTable t = new DataTable(sTab);
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "tmp_ecr",
                Caption = "ECR",
                MaxLength = 9,
                ReadOnly = false,
                DefaultValue = (String)""
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "tmp_ecd",
                Caption = "Descrizione",
                MaxLength = 50,
                ReadOnly = false,
                DefaultValue = (String)""
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "tmp_rep",
                Caption = "Reparto",
                MaxLength = 3,
                ReadOnly = false,
                DefaultValue = (String)""
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "tmp_red",
                Caption = "Descrizione",
                MaxLength = 50,
                ReadOnly = false,
                DefaultValue = (String)""
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "tmp_art",
                Caption = "Codice articolo",
                MaxLength = 7,
                ReadOnly = false,
                DefaultValue = (String)""
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "tmp_ard",
                Caption = "Descrizione",
                MaxLength = 50,
                ReadOnly = false,
                DefaultValue = (String)""
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "tmp_umi",
                Caption = "UM",
                MaxLength = 2,
                ReadOnly = false,
                DefaultValue = (String)""
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.Decimal"),
                ColumnName = "tmp_iro",
                Caption = "Indice di rotazione",
                ReadOnly = false,
                DefaultValue = (Decimal)0
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.Decimal"),
                ColumnName = "tmp_scm",
                Caption = "Scorta media",
                ReadOnly = false,
                DefaultValue = (Decimal)0
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.Decimal"),
                ColumnName = "tmp_cve",
                Caption = "Costo del venduto",
                ReadOnly = false,
                DefaultValue = (Decimal)0
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.Decimal"),
                ColumnName = "inv_000",
                Caption = "Inventario",
                ReadOnly = false,
                DefaultValue = (Decimal)0
            });

            for (int i = intIni; i <= intFin; i++)
            {
                string sPer = i.ToString("000");

                t.Columns.Add(new DataColumn()
                {
                    DataType = Type.GetType("System.Decimal"),
                    ColumnName = "acq_" + sPer,
                    Caption = "Acquistato",
                    ReadOnly = false,
                    DefaultValue = (Decimal)0
                });
                t.Columns.Add(new DataColumn()
                {
                    DataType = Type.GetType("System.Decimal"),
                    ColumnName = "ven_" + sPer,
                    Caption = "Vanduto",
                    ReadOnly = false,
                    DefaultValue = (Decimal)0
                });
                //t.Columns.Add(new DataColumn()
                //{
                //    DataType = Type.GetType("System.Decimal"),
                //    ColumnName = "gia_" + sPer,
                //    Caption = "Giacenza",
                //    ReadOnly = false,
                //    DefaultValue = (Decimal)0
                //});
                //t.Columns.Add(new DataColumn()
                //{
                //    DataType = Type.GetType("System.Decimal"),
                //    ColumnName = "iro_" + sPer,
                //    Caption = "Indice rotazione",
                //    ReadOnly = false,
                //    DefaultValue = (Decimal)0
                //});
            }

            return t;
        }

        public DataTable TabTmpArtMovimenti(string sTab)
        {
            DataTable t = new DataTable(sTab);
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "tmp_ecr",
                Caption = "ECR",
                MaxLength = 9,
                ReadOnly = false,
                DefaultValue = (String)""
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "tmp_ecd",
                Caption = "Descrizione",
                MaxLength = 50,
                ReadOnly = false,
                DefaultValue = (String)""
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "tmp_rep",
                Caption = "Reparto",
                MaxLength = 3,
                ReadOnly = false,
                DefaultValue = (String)""
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "tmp_red",
                Caption = "Descrizione",
                MaxLength = 50,
                ReadOnly = false,
                DefaultValue = (String)""
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "tmp_art",
                Caption = "Codice articolo",
                MaxLength = 7,
                ReadOnly = false,
                DefaultValue = (String)""
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "tmp_ard",
                Caption = "Descrizione",
                MaxLength = 50,
                ReadOnly = false,
                DefaultValue = (String)""
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "tmp_umi",
                Caption = "UM",
                MaxLength = 2,
                ReadOnly = false,
                DefaultValue = (String)""
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.Decimal"),
                ColumnName = "tmp_acq",
                Caption = "Acquistato",
                ReadOnly = false,
                DefaultValue = (Decimal)0
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.Decimal"),
                ColumnName = "tmp_ven",
                Caption = "Venduto",
                ReadOnly = false,
                DefaultValue = (Decimal)0
            });

            return t;

        }

        public DataTable TabTmpBil(string sTab)
        {
            DataTable t = new DataTable(sTab);

            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "bil_ban",
                Caption = "Banco",
                MaxLength = 1,
                ReadOnly = false,
                DefaultValue = (String)""
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "bil_plu",
                Caption = "Codice",
                MaxLength = 4,
                ReadOnly = false,
                DefaultValue = (String)""
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.Decimal"),
                ColumnName = "bil_qta",
                Caption = "Qta",
                ReadOnly = false,
                DefaultValue = (Decimal)0
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.Decimal"),
                ColumnName = "bil_prv",
                Caption = "Prezzo al kg",
                ReadOnly = false,
                DefaultValue = (Decimal)0
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.Decimal"),
                ColumnName = "bil_imp",
                Caption = "Imp",
                ReadOnly = false,
                DefaultValue = (Decimal)0
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "bil_ims",
                Caption = "Importo in stringa",
                MaxLength = 8,
                ReadOnly = false,
                DefaultValue = (String)""
            });

            return t;
        }

        public DataTable TabTmpStaStorico(string sTab)
        {
            DataTable t = new DataTable(sTab);

            if (sTab == "TmpArt")
            {
                t.Columns.Add(new DataColumn()
                {
                    DataType = Type.GetType("System.String"),
                    ColumnName = "sta_art",
                    Caption = "Codice",
                    MaxLength = 7,
                    ReadOnly = false,
                    DefaultValue = (object)""
                });
                t.Columns.Add(new DataColumn()
                {
                    DataType = Type.GetType("System.String"),
                    ColumnName = "sta_des",
                    Caption = "Descrizione",
                    MaxLength = 50,
                    ReadOnly = false,
                    DefaultValue = (object)""
                });
                t.Columns.Add(new DataColumn()
                {
                    DataType = Type.GetType("System.String"),
                    ColumnName = "art_umi",
                    Caption = "UMI",
                    MaxLength = 2,
                    ReadOnly = false,
                    DefaultValue = (object)""
                });
                t.Columns.Add(new DataColumn()
                {
                    DataType = Type.GetType("System.String"),
                    ColumnName = "art_rep",
                    Caption = "Reparto",
                    MaxLength = 3,
                    ReadOnly = false,
                    DefaultValue = (object)""
                });
                t.Columns.Add(new DataColumn()
                {
                    DataType = Type.GetType("System.String"),
                    ColumnName = "art_red",
                    Caption = "Reparto",
                    MaxLength = 30,
                    ReadOnly = false,
                    DefaultValue = (object)""
                });
                t.Columns.Add(new DataColumn()
                {
                    DataType = Type.GetType("System.String"),
                    ColumnName = "art_ecr",
                    Caption = "ECR",
                    MaxLength = 9,
                    ReadOnly = false,
                    DefaultValue = (object)""
                });
                t.Columns.Add(new DataColumn()
                {
                    DataType = Type.GetType("System.String"),
                    ColumnName = "art_ecd",
                    Caption = "Ecr",
                    MaxLength = 30,
                    ReadOnly = false,
                    DefaultValue = (object)""
                });
            }
            else
            {
                t.Columns.Add(new DataColumn()
                {
                    DataType = Type.GetType("System.String"),
                    ColumnName = "str_rep",
                    Caption = "Codice",
                    MaxLength = 3,
                    ReadOnly = false,
                    DefaultValue = (object)""
                });
                t.Columns.Add(new DataColumn()
                {
                    DataType = Type.GetType("System.String"),
                    ColumnName = "str_des",
                    Caption = "Descrizione",
                    MaxLength = 50,
                    ReadOnly = false,
                    DefaultValue = (object)""
                });
            }

            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.Decimal"),
                ColumnName = "acq_qta",
                Caption = "Qta acquisto",
                ReadOnly = false,
                DefaultValue = (decimal)0
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.Decimal"),
                ColumnName = "acq_qkg",
                Caption = "Qta acquisto",
                ReadOnly = false,
                DefaultValue = (decimal)0
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.Decimal"),
                ColumnName = "acq_val",
                Caption = "Val acquisto",
                ReadOnly = false,
                DefaultValue = (decimal)0
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.Decimal"),
                ColumnName = "ven_qta",
                Caption = "Qta venduto",
                ReadOnly = false,
                DefaultValue = (decimal)0
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.Decimal"),
                ColumnName = "ven_qkg",
                Caption = "Qta venduto",
                ReadOnly = false,
                DefaultValue = (decimal)0
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.Decimal"),
                ColumnName = "ven_val",
                Caption = "Val venduto",
                ReadOnly = false,
                DefaultValue = (decimal)0
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.Decimal"),
                ColumnName = "ven_vni",
                Caption = "Val venduto netto IVA",
                ReadOnly = false,
                DefaultValue = (decimal)0
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.Decimal"),
                ColumnName = "ven_vcs",
                Caption = "Val venduto al costo se presente",
                ReadOnly = false,
                DefaultValue = (decimal)0
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.Decimal"),
                ColumnName = "mov_qta",
                Caption = "Mov venduto",
                ReadOnly = false,
                DefaultValue = (decimal)0
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.Decimal"),
                ColumnName = "mov_qkg",
                Caption = "Mov venduto",
                ReadOnly = false,
                DefaultValue = (decimal)0
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.Decimal"),
                ColumnName = "mov_val",
                Caption = "Mov venduto",
                ReadOnly = false,
                DefaultValue = (decimal)0
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.Decimal"),
                ColumnName = "off_qta",
                Caption = "Off venduto",
                ReadOnly = false,
                DefaultValue = (decimal)0
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.Decimal"),
                ColumnName = "off_qkg",
                Caption = "Off venduto",
                ReadOnly = false,
                DefaultValue = (decimal)0
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.Decimal"),
                ColumnName = "off_val",
                Caption = "Off venduto",
                ReadOnly = false,
                DefaultValue = (decimal)0
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.Decimal"),
                ColumnName = "fid_qta",
                Caption = "Fid venduto",
                ReadOnly = false,
                DefaultValue = (decimal)0
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.Decimal"),
                ColumnName = "fid_qkg",
                Caption = "Fid venduto",
                ReadOnly = false,
                DefaultValue = (decimal)0
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.Decimal"),
                ColumnName = "fid_val",
                Caption = "Fid venduto",
                ReadOnly = false,
                DefaultValue = (decimal)0
            });

            return t;
        }

        public DataTable TabTmpPosCassieri(string strTab, string strUsr)
        {
            DataTable t = new DataTable(strTab);
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "TmpTre",
                Caption = "Tipo record",
                MaxLength = 2,
                ReadOnly = false,
                DefaultValue = (String)""
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "TmpCod",
                Caption = "Codice riga",
                MaxLength = 3,
                ReadOnly = false,
                DefaultValue = (String)""
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "TmpDes",
                Caption = "Descrizione",
                MaxLength = 30,
                ReadOnly = false,
                DefaultValue = (String)""
            });

            string[] a = strUsr.Split(';');

            foreach (string ss in a)
            {
                if (ss != "")
                {
                    t.Columns.Add(new DataColumn()
                    {
                        DataType = Type.GetType("System.Decimal"),
                        ColumnName = ss,
                        Caption = "Valore",
                        ReadOnly = false,
                        DefaultValue = (Decimal)0
                    });
                }
            }

            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.Decimal"),
                ColumnName = "TmpTot",
                Caption = "Totali",
                ReadOnly = false,
                DefaultValue = (Decimal)0
            });

            return t;
        }

        public DataTable TabTmpStaSettimana(string strTab)
        {
            DataTable t = new DataTable(strTab);
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "TmpDay",
                Caption = "Data",
                MaxLength = 8,
                ReadOnly = false,
                DefaultValue = (String)""
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "TmpRep",
                Caption = "Reparto",
                MaxLength = 3,
                ReadOnly = false,
                DefaultValue = (String)""
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "TmpRed",
                Caption = "Reparto",
                MaxLength = 30,
                ReadOnly = false,
                DefaultValue = (String)""
            });

            for (int i = 0; i <= 7; i++)
            {
                t.Columns.Add(new DataColumn()
                {
                    DataType = Type.GetType("System.Decimal"),
                    ColumnName = "TmpD" + i.ToString() + "v",
                    Caption = "Venduto",
                    ReadOnly = false,
                    DefaultValue = (Decimal)0
                });
                t.Columns.Add(new DataColumn()
                {
                    DataType = Type.GetType("System.Decimal"),
                    ColumnName = "TmpD" + i.ToString() + "i",
                    Caption = "Venduto netto IVA con costo",
                    ReadOnly = false,
                    DefaultValue = (Decimal)0
                });
                t.Columns.Add(new DataColumn()
                {
                    DataType = Type.GetType("System.Decimal"),
                    ColumnName = "TmpD" + i.ToString() + "a",
                    Caption = "Acquisti",
                    ReadOnly = false,
                    DefaultValue = (Decimal)0
                });
                t.Columns.Add(new DataColumn()
                {
                    DataType = Type.GetType("System.Decimal"),
                    ColumnName = "TmpD" + i.ToString() + "m",
                    Caption = "Margine",
                    ReadOnly = false,
                    DefaultValue = (Decimal)0
                });
                t.Columns.Add(new DataColumn()
                {
                    DataType = Type.GetType("System.Decimal"),
                    ColumnName = "TmpD" + i.ToString() + "c",
                    Caption = "N.clienti",
                    ReadOnly = false,
                    DefaultValue = (Decimal)0
                });
            }

            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.Decimal"),
                ColumnName = "TmpInc",
                Caption = "Incidenza",
                ReadOnly = false,
                DefaultValue = (Decimal)0
            });

            return t;
        }

        public DataTable TabTmpStaSettimanaFascia(string strTab)
        {
            DataTable t = new DataTable(strTab);
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "TmpDay",
                Caption = "Data",
                MaxLength = 8,
                ReadOnly = false,
                DefaultValue = (String)""
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "TmpOra",
                Caption = "Orario",
                MaxLength = 5,
                ReadOnly = false,
                DefaultValue = (String)""
            });

            for (int i = 0; i <= 7; i++)
            {
                t.Columns.Add(new DataColumn()
                {
                    DataType = Type.GetType("System.Decimal"),
                    ColumnName = "TmpD" + i.ToString() + "v",
                    Caption = "Venduto",
                    ReadOnly = false,
                    DefaultValue = (Decimal)0
                });
                t.Columns.Add(new DataColumn()
                {
                    DataType = Type.GetType("System.Decimal"),
                    ColumnName = "TmpD" + i.ToString() + "c",
                    Caption = "N.clienti",
                    ReadOnly = false,
                    DefaultValue = (Decimal)0
                });
                t.Columns.Add(new DataColumn()
                {
                    DataType = Type.GetType("System.Decimal"),
                    ColumnName = "TmpD" + i.ToString() + "m",
                    Caption = "Margine",
                    ReadOnly = false,
                    DefaultValue = (Decimal)0
                });
            }

            return t;
        }

        public DataTable TabTmpSta(string strTab)
        {
            DataTable t = new DataTable(strTab);
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "tab_cod",
                Caption = "Codice",
                MaxLength = 3,
                ReadOnly = false,
                DefaultValue = (String)""
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "tab_des",
                Caption = "Descrizione",
                MaxLength = 20,
                ReadOnly = false,
                DefaultValue = (String)""
            });

            return t;
        }

        public DataTable TabTmpTermDett(string strTab)
        {
            DataTable t = new DataTable(strTab);
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "trm_ean",
                Caption = "Ean",
                MaxLength = 13,
                ReadOnly = false,
                DefaultValue = (String)""
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "trm_art",
                Caption = "Articolo",
                MaxLength = 7,
                ReadOnly = false,
                DefaultValue = (String)""
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "trm_ard",
                Caption = "Articolo",
                MaxLength = 50,
                ReadOnly = false,
                DefaultValue = (String)""
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.Decimal"),
                ColumnName = "trm_qta",
                Caption = "Q.tà",
                ReadOnly = false,
                DefaultValue = (Decimal)0
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "trm_tmp",
                Caption = "Riga",
                MaxLength = 100,
                ReadOnly = false,
                DefaultValue = (String)""
            });

            return t;
        }

        public DataTable TabTmpArtOff(string sTab)
        {
            DataTable t = new DataTable(sTab);
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "tmp_art",
                Caption = "Art",
                MaxLength = 7,
                ReadOnly = false,
                DefaultValue = (String)""
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "tmp_ard",
                Caption = "Descrizione",
                MaxLength = 50,
                ReadOnly = false,
                DefaultValue = (String)""
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "tmp_arb",
                Caption = "Descrizione breve",
                MaxLength = 25,
                ReadOnly = false,
                DefaultValue = (String)""
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "tmp_sta",
                Caption = "Stato",
                MaxLength = 1,
                ReadOnly = false,
                DefaultValue = (String)""
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "tmp_iva",
                Caption = "IVA",
                MaxLength = 3,
                ReadOnly = false,
                DefaultValue = (String)""
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "tmp_umi",
                Caption = "UMI",
                MaxLength = 2,
                ReadOnly = false,
                DefaultValue = (String)""
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "tmp_rep",
                Caption = "Reparto",
                MaxLength = 3,
                ReadOnly = false,
                DefaultValue = (String)""
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "tmp_ecr",
                Caption = "Merceologia",
                MaxLength = 9,
                ReadOnly = false,
                DefaultValue = (String)""
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "tmp_for",
                Caption = "Fornitore",
                MaxLength = 5,
                ReadOnly = false,
                DefaultValue = (String)""
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "tmp_fod",
                Caption = "Fornitore",
                MaxLength = 70,
                ReadOnly = false,
                DefaultValue = (String)""
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "tmp_arf",
                Caption = "Art. fornitore",
                MaxLength = 20,
                ReadOnly = false,
                DefaultValue = (String)""
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "tmp_lis",
                Caption = "Listino",
                MaxLength = 3,
                ReadOnly = false,
                DefaultValue = (String)""
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "tmp_liv",
                Caption = "Listino ivato S/N",
                MaxLength = 1,
                ReadOnly = false,
                DefaultValue = (String)""
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "tmp_lva",
                Caption = "Listino vendita annullato",
                MaxLength = 3,
                ReadOnly = false,
                DefaultValue = (String)""
            });
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
                DataType = Type.GetType("System.Decimal"),
                ColumnName = "tmp_prp",
                Caption = "Vendita al pubblico",
                ReadOnly = false,
                DefaultValue = (Decimal)0
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.Decimal"),
                ColumnName = "tmp_ppv",
                Caption = "Prezzo vendita pieno in caso di lis=PRO",
                ReadOnly = false,
                DefaultValue = (Decimal)0
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.Boolean"),
                ColumnName = "tmp_bil",
                Caption = "Bilancia",
                ReadOnly = false,
                DefaultValue = (Boolean)false
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "tmp_tgr",
                Caption = "T.gramm",
                MaxLength = 2,
                ReadOnly = false,
                DefaultValue = (String)""
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.Decimal"),
                ColumnName = "tmp_pxc",
                Caption = "Pezzi",
                ReadOnly = false,
                DefaultValue = (Decimal)0
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.Decimal"),
                ColumnName = "tmp_pne",
                Caption = "Peso netto",
                ReadOnly = false,
                DefaultValue = (Decimal)0
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.Decimal"),
                ColumnName = "tmp_sfr",
                Caption = "Sfrido",
                ReadOnly = false,
                DefaultValue = (Decimal)0
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.Decimal"),
                ColumnName = "tmp_tar",
                Caption = "Tara",
                ReadOnly = false,
                DefaultValue = (Decimal)0
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "tmp_ori",
                Caption = "Origine",
                MaxLength = 3,
                ReadOnly = false,
                DefaultValue = (String)""
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "tmp_cal",
                Caption = "Calibro",
                MaxLength = 3,
                ReadOnly = false,
                DefaultValue = (String)""
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "tmp_cat",
                Caption = "Categoria",
                MaxLength = 3,
                ReadOnly = false,
                DefaultValue = (String)""
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "tmp_reb",
                Caption = "Rep bilancia",
                MaxLength = 1,
                ReadOnly = false,
                DefaultValue = (String)""
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "tmp_gsc",
                Caption = "GG scad",
                MaxLength = 3,
                ReadOnly = false,
                DefaultValue = (String)""
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "tmp_plu",
                Caption = "PLU",
                MaxLength = 6,
                ReadOnly = false,
                DefaultValue = (String)""
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "tmp_tas",
                Caption = "Tasto bilancia",
                MaxLength = 4,
                ReadOnly = false,
                DefaultValue = (String)""
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "tmp_tra",
                Caption = "Tracciabilità",
                MaxLength = 3,
                ReadOnly = false,
                DefaultValue = (String)""
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.Boolean"),
                ColumnName = "tmp_bpz",
                Caption = "A corpo",
                ReadOnly = false,
                DefaultValue = (Boolean)false
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
                DataType = Type.GetType("System.Decimal"),
                ColumnName = "tmp_eaq",
                Caption = "Q.tà x ean",
                ReadOnly = false,
                DefaultValue = (Decimal)0
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.Decimal"),
                ColumnName = "tmp_eap",
                Caption = "Prezzo per ean",
                ReadOnly = false,
                DefaultValue = (Decimal)0
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.Boolean"),
                ColumnName = "tmp_ecp",
                Caption = "Ean con peso",
                ReadOnly = false,
                DefaultValue = (Boolean)false
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.Boolean"),
                ColumnName = "tmp_eaa",
                Caption = "Ean annullato",
                ReadOnly = false,
                DefaultValue = (Boolean)false
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.Decimal"),
                ColumnName = "tmp_epu",
                Caption = "Punti per ean",
                ReadOnly = false,
                DefaultValue = (Decimal)0
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "tmp_cam",
                Caption = "Campagna",
                MaxLength = 3,
                ReadOnly = false,
                DefaultValue = (String)""
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "tmp_eti",
                Caption = "Etichetta",
                MaxLength = 3,
                ReadOnly = false,
                DefaultValue = (String)""
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "tmp_eqp",
                Caption = "Equivalenza prezzo",
                MaxLength = 4,
                ReadOnly = false,
                DefaultValue = (String)""
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "tmp_lot",
                Caption = "Lotto da barcode",
                MaxLength = 15,
                ReadOnly = false,
                DefaultValue = (String)""
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "tmp_img",
                Caption = "Immagine",
                MaxLength = 50,
                ReadOnly = false,
                DefaultValue = (String)""
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.DateTime"),
                ColumnName = "oft_dti",
                Caption = "Offerta inizio",
                ReadOnly = false
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.DateTime"),
                ColumnName = "oft_dtf",
                Caption = "Offerta fine",
                ReadOnly = false
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "oft_yea",
                Caption = "Anno",
                MaxLength = 4,
                ReadOnly = false,
                DefaultValue = (String)""
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "oft_cod",
                Caption = "Codice",
                MaxLength = 3,
                ReadOnly = false,
                DefaultValue = (String)""
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "oft_des",
                Caption = "Descrizione",
                MaxLength = 50,
                ReadOnly = false,
                DefaultValue = (String)""
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "ofa_rig",
                Caption = "Riga",
                MaxLength = 4,
                ReadOnly = false,
                DefaultValue = (String)""
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "ofa_tip",
                Caption = "Tipo",
                MaxLength = 3,
                ReadOnly = false,
                DefaultValue = (String)""
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.Decimal"),
                ColumnName = "ofa_val",
                Caption = "Valore",
                ReadOnly = false,
                DefaultValue = (Decimal)0
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.Decimal"),
                ColumnName = "ofa_xem",
                Caption = "Valore",
                ReadOnly = false,
                DefaultValue = (Decimal)0
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.Decimal"),
                ColumnName = "ofa_xen",
                Caption = "Valore",
                ReadOnly = false,
                DefaultValue = (Decimal)0
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.Decimal"),
                ColumnName = "ofa_mix",
                Caption = "Valore",
                ReadOnly = false,
                DefaultValue = (Decimal)0
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.Decimal"),
                ColumnName = "ofa_pun",
                Caption = "Valore",
                ReadOnly = false,
                DefaultValue = (Decimal)0
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.Boolean"),
                ColumnName = "ofa_ann",
                Caption = "Valore",
                ReadOnly = false,
                DefaultValue = (Boolean)false
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "ofa_not",
                Caption = "Note",
                MaxLength = 50,
                ReadOnly = false,
                DefaultValue = (String)""
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.Decimal"),
                ColumnName = "tmp_gia",
                Caption = "Giacenza",
                ReadOnly = false,
                DefaultValue = (Decimal)0
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "tmp_tof",
                Caption = "Tipo offerta",
                MaxLength = 3,
                ReadOnly = false,
                DefaultValue = (String)""
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.Decimal"),
                ColumnName = "tmp_pro",
                Caption = "Prezzo Offerta",
                ReadOnly = false,
                DefaultValue = (Decimal)0
            });

            return t;
        }

        public DataTable TabTmpPrzNeg(string sTab)
        {
            DataTable t = new DataTable(sTab);

            t = new DataTable(sTab)
            {
                Columns = {
                    new DataColumn()
                    {
                        DataType = Type.GetType("System.String"),
                        ColumnName = "art_rep",
                        Caption = "Reparto",
                        MaxLength = 3,
                        ReadOnly = false,
                        DefaultValue = (String)""
                    },
                    new DataColumn()
                    {
                        DataType = Type.GetType("System.String"),
                        ColumnName = "art_red",
                        Caption = "Rep",
                        MaxLength = 30,
                        ReadOnly = false,
                        DefaultValue = (String)""
                    },
                    new DataColumn()
                    {
                        DataType = Type.GetType("System.String"),
                        ColumnName = "art_cod",
                        Caption = "Art",
                        MaxLength = 7,
                        ReadOnly = false,
                        DefaultValue = (String)""
                    },
                    new DataColumn()
                    {
                        DataType = Type.GetType("System.String"),
                        ColumnName = "art_des",
                        Caption = "Des",
                        MaxLength = 50,
                        ReadOnly = false,
                        DefaultValue = (String)""
                    },
                    new DataColumn()
                    {
                        DataType = Type.GetType("System.String"),
                        ColumnName = "art_umi",
                        Caption = "UM",
                        MaxLength = 2,
                        ReadOnly = false,
                        DefaultValue = (String)""
                    },
                    new DataColumn()
                    {
                        DataType = Type.GetType("System.DateTime"),
                        ColumnName = "001_day",
                        Caption = "Data prezzo negozio",
                        ReadOnly = false,
                    },
                    new DataColumn()
                    {
                        DataType = Type.GetType("System.Decimal"),
                        ColumnName = "001_prv",
                        Caption = "Prezzo negozio",
                        ReadOnly = false,
                        DefaultValue = (Decimal)0
                    },
                    new DataColumn()
                    {
                        DataType = Type.GetType("System.DateTime"),
                        ColumnName = "002_day",
                        Caption = "Data prezzo negozio",
                        ReadOnly = false,
                    },
                    new DataColumn()
                    {
                        DataType = Type.GetType("System.Decimal"),
                        ColumnName = "002_prv",
                        Caption = "Prezzo negozio",
                        ReadOnly = false,
                        DefaultValue = (Decimal)0
                    },
                    new DataColumn()
                    {
                        DataType = Type.GetType("System.Decimal"),
                        ColumnName = "art_gia",
                        Caption = "Giacenza",
                        ReadOnly = false,
                        DefaultValue = (Decimal)0
                    },
                }
            };

            return t;
        }

    }
}
