using System;
using System.Data;
using System.Data.SqlClient;

namespace APOffice
{
    public class clsDbMigration
    {
        private clsFuncs _clsFun = new clsFuncs();

        public void CheckDb(string strConSql)
        {
            try
            {
                MigraV001_TabEtiFormatiLayout(strConSql);
                MigraV002_TabEtiFormatiDefault(strConSql);
                MigraV003_AnaArticoli_art_img(strConSql);
                MigraV004_TabEtiFormatiLayout_ImgLengths(strConSql);
                MigraV005_TabParametri_Completa(strConSql);
                MigraV006_Indexes_Optimization(strConSql);
                MigraV007_TabEti_StatoPostStampa(strConSql);
                MigraV008_TabEti_TestoLibero(strConSql);
                MigraV009_CleanPlaceholderAndScaleOptimization(strConSql);
            }
            catch (Exception ex)
            {
                Console.WriteLine("Errore durante l'auto-migrazione DB: " + ex.Message);
            }
        }

        private void MigraV001_TabEtiFormatiLayout(string strConSql)
        {
            // Verifica se la tabella TabEtiFormatiLayout esiste
            string s = "IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[TabEtiFormatiLayout]') AND type in (N'U')) " +
                       "BEGIN " +
                       "CREATE TABLE [dbo].[TabEtiFormatiLayout]( " +
                       "[eti_cod] [nvarchar](3) NOT NULL, " +
                       "[eti_des] [nvarchar](50) NOT NULL, " +
                       "[eti_tip] [nvarchar](3) NOT NULL CONSTRAINT [DF_TabEtiFormatiLayout_eti_tip] DEFAULT ('GDI'), " +
                       "[eti_pri] [nvarchar](50) NOT NULL CONSTRAINT [DF_TabEtiFormatiLayout_eti_pri] DEFAULT (''), " +
                       "[eti_dim] [nvarchar](20) NOT NULL CONSTRAINT [DF_TabEtiFormatiLayout_eti_dim] DEFAULT ('300,200'), " +
                       "[eti_imm] [nvarchar](1) NOT NULL CONSTRAINT [DF_TabEtiFormatiLayout_eti_imm] DEFAULT ('N'), " +
                       "[eti_s003] [nvarchar](200) NOT NULL CONSTRAINT [DF_TabEtiFormatiLayout_eti_s003] DEFAULT (''), " +
                       "[eti_s004] [nvarchar](200) NOT NULL CONSTRAINT [DF_TabEtiFormatiLayout_eti_s004] DEFAULT (''), " +
                       "[eti_s005] [nvarchar](200) NOT NULL CONSTRAINT [DF_TabEtiFormatiLayout_eti_s005] DEFAULT (''), " +
                       "[eti_s006] [nvarchar](200) NOT NULL CONSTRAINT [DF_TabEtiFormatiLayout_eti_s006] DEFAULT (''), " +
                       "[eti_s007] [nvarchar](200) NOT NULL CONSTRAINT [DF_TabEtiFormatiLayout_eti_s007] DEFAULT (''), " +
                       "[eti_s008] [nvarchar](200) NOT NULL CONSTRAINT [DF_TabEtiFormatiLayout_eti_s008] DEFAULT (''), " +
                       "[eti_s009] [nvarchar](200) NOT NULL CONSTRAINT [DF_TabEtiFormatiLayout_eti_s009] DEFAULT (''), " +
                       "[eti_s010] [nvarchar](200) NOT NULL CONSTRAINT [DF_TabEtiFormatiLayout_eti_s010] DEFAULT (''), " +
                       "[eti_s011] [nvarchar](200) NOT NULL CONSTRAINT [DF_TabEtiFormatiLayout_eti_s011] DEFAULT (''), " +
                       "[eti_s012] [nvarchar](200) NOT NULL CONSTRAINT [DF_TabEtiFormatiLayout_eti_s012] DEFAULT (''), " +
                       "[eti_s013] [nvarchar](200) NOT NULL CONSTRAINT [DF_TabEtiFormatiLayout_eti_s013] DEFAULT (''), " +
                       "[eti_s014] [nvarchar](200) NOT NULL CONSTRAINT [DF_TabEtiFormatiLayout_eti_s014] DEFAULT (''), " +
                       "[eti_s015] [nvarchar](200) NOT NULL CONSTRAINT [DF_TabEtiFormatiLayout_eti_s015] DEFAULT (''), " +
                       "[eti_s016] [nvarchar](200) NOT NULL CONSTRAINT [DF_TabEtiFormatiLayout_eti_s016] DEFAULT (''), " +
                       "[eti_s021] [nvarchar](200) NOT NULL CONSTRAINT [DF_TabEtiFormatiLayout_eti_s021] DEFAULT (''), " +
                       "[eti_s022] [nvarchar](200) NOT NULL CONSTRAINT [DF_TabEtiFormatiLayout_eti_s022] DEFAULT (''), " +
                       "[eti_t001] [nvarchar](500) NOT NULL CONSTRAINT [DF_TabEtiFormatiLayout_eti_t001] DEFAULT (''), " +
                       "[eti_t002] [nvarchar](500) NOT NULL CONSTRAINT [DF_TabEtiFormatiLayout_eti_t002] DEFAULT (''), " +
                       "[eti_t003] [nvarchar](500) NOT NULL CONSTRAINT [DF_TabEtiFormatiLayout_eti_t003] DEFAULT (''), " +
                       "[eti_t004] [nvarchar](500) NOT NULL CONSTRAINT [DF_TabEtiFormatiLayout_eti_t004] DEFAULT (''), " +
                       "[eti_t005] [nvarchar](500) NOT NULL CONSTRAINT [DF_TabEtiFormatiLayout_eti_t005] DEFAULT (''), " +
                       "[eti_t006] [nvarchar](500) NOT NULL CONSTRAINT [DF_TabEtiFormatiLayout_eti_t006] DEFAULT (''), " +
                       "[eti_t007] [nvarchar](500) NOT NULL CONSTRAINT [DF_TabEtiFormatiLayout_eti_t007] DEFAULT (''), " +
                       "[eti_t008] [nvarchar](500) NOT NULL CONSTRAINT [DF_TabEtiFormatiLayout_eti_t008] DEFAULT (''), " +
                       "[eti_h001] [nvarchar](200) NOT NULL CONSTRAINT [DF_TabEtiFormatiLayout_eti_h001] DEFAULT (''), " +
                       "[eti_h002] [nvarchar](200) NOT NULL CONSTRAINT [DF_TabEtiFormatiLayout_eti_h002] DEFAULT (''), " +
                       "[eti_h003] [nvarchar](200) NOT NULL CONSTRAINT [DF_TabEtiFormatiLayout_eti_h003] DEFAULT (''), " +
                       "[eti_h004] [nvarchar](200) NOT NULL CONSTRAINT [DF_TabEtiFormatiLayout_eti_h004] DEFAULT (''), " +
                       "[eti_h005] [nvarchar](200) NOT NULL CONSTRAINT [DF_TabEtiFormatiLayout_eti_h005] DEFAULT (''), " +
                       "[eti_h006] [nvarchar](200) NOT NULL CONSTRAINT [DF_TabEtiFormatiLayout_eti_h006] DEFAULT (''), " +
                       "[eti_h007] [nvarchar](200) NOT NULL CONSTRAINT [DF_TabEtiFormatiLayout_eti_h007] DEFAULT (''), " +
                       "[eti_h008] [nvarchar](200) NOT NULL CONSTRAINT [DF_TabEtiFormatiLayout_eti_h008] DEFAULT (''), " +
                       "[eti_pag] [nvarchar](50) NOT NULL CONSTRAINT [DF_TabEtiFormatiLayout_eti_pag] DEFAULT (''), " +
                       "[eti_ann] [bit] NOT NULL CONSTRAINT [DF_TabEtiFormatiLayout_eti_ann] DEFAULT ((0)), " +
                       "CONSTRAINT [PK_TabEtiFormatiLayout] PRIMARY KEY CLUSTERED ([eti_cod] ASC) " +
                       ") ON [PRIMARY] " +
                       "END " +
                       "ELSE BEGIN " +
                       "IF COL_LENGTH('TabEtiFormatiLayout', 'eti_pag') IS NULL " +
                       "ALTER TABLE TabEtiFormatiLayout ADD eti_pag nvarchar(50) NOT NULL CONSTRAINT DF_TabEtiFormatiLayout_eti_pag DEFAULT (''); " +
                       "IF COL_LENGTH('TabEtiFormatiLayout', 'eti_s014') IS NULL " +
                       "ALTER TABLE TabEtiFormatiLayout ADD eti_s014 nvarchar(200) NOT NULL CONSTRAINT DF_TabEtiFormatiLayout_eti_s014 DEFAULT (''); " +
                       "IF COL_LENGTH('TabEtiFormatiLayout', 'eti_s015') IS NULL " +
                       "ALTER TABLE TabEtiFormatiLayout ADD eti_s015 nvarchar(200) NOT NULL CONSTRAINT DF_TabEtiFormatiLayout_eti_s015 DEFAULT (''); " +
                       "IF COL_LENGTH('TabEtiFormatiLayout', 'eti_s016') IS NULL " +
                       "ALTER TABLE TabEtiFormatiLayout ADD eti_s016 nvarchar(200) NOT NULL CONSTRAINT DF_TabEtiFormatiLayout_eti_s016 DEFAULT (''); " +
                       "IF COL_LENGTH('TabEtiFormatiLayout', 'eti_s017') IS NULL " +
                       "ALTER TABLE TabEtiFormatiLayout ADD eti_s017 nvarchar(200) NOT NULL CONSTRAINT DF_TabEtiFormatiLayout_eti_s017 DEFAULT (''); " +
                       "IF COL_LENGTH('TabEtiFormatiLayout', 'eti_s018') IS NULL " +
                       "ALTER TABLE TabEtiFormatiLayout ADD eti_s018 nvarchar(200) NOT NULL CONSTRAINT DF_TabEtiFormatiLayout_eti_s018 DEFAULT (''); " +
                       "IF COL_LENGTH('TabEtiFormatiLayout', 'eti_s019') IS NULL " +
                       "ALTER TABLE TabEtiFormatiLayout ADD eti_s019 nvarchar(200) NOT NULL CONSTRAINT DF_TabEtiFormatiLayout_eti_s019 DEFAULT (''); " +
                       "IF COL_LENGTH('TabEtiFormatiLayout', 'eti_s020') IS NULL " +
                       "ALTER TABLE TabEtiFormatiLayout ADD eti_s020 nvarchar(200) NOT NULL CONSTRAINT DF_TabEtiFormatiLayout_eti_s020 DEFAULT (''); " +
                       "IF COL_LENGTH('TabEtiFormatiLayout', 'eti_s021') IS NULL " +
                       "ALTER TABLE TabEtiFormatiLayout ADD eti_s021 nvarchar(200) NOT NULL CONSTRAINT DF_TabEtiFormatiLayout_eti_s021 DEFAULT (''); " +
                       "IF COL_LENGTH('TabEtiFormatiLayout', 'eti_s022') IS NULL " +
                       "ALTER TABLE TabEtiFormatiLayout ADD eti_s022 nvarchar(200) NOT NULL CONSTRAINT DF_TabEtiFormatiLayout_eti_s022 DEFAULT (''); " +
                       "IF COL_LENGTH('TabEtiFormatiLayout', 'eti_s023') IS NULL " +
                       "ALTER TABLE TabEtiFormatiLayout ADD eti_s023 nvarchar(200) NOT NULL CONSTRAINT DF_TabEtiFormatiLayout_eti_s023 DEFAULT (''); " +
                       "IF COL_LENGTH('TabEtiFormatiLayout', 'eti_s024') IS NULL " +
                       "ALTER TABLE TabEtiFormatiLayout ADD eti_s024 nvarchar(200) NOT NULL CONSTRAINT DF_TabEtiFormatiLayout_eti_s024 DEFAULT (''); " +
                       "IF COL_LENGTH('TabEtiFormatiLayout', 'eti_s025') IS NULL " +
                       "ALTER TABLE TabEtiFormatiLayout ADD eti_s025 nvarchar(200) NOT NULL CONSTRAINT DF_TabEtiFormatiLayout_eti_s025 DEFAULT (''); " +
                       "IF COL_LENGTH('TabEtiFormatiLayout', 'eti_t001') IS NULL " +
                       "ALTER TABLE TabEtiFormatiLayout ADD eti_t001 nvarchar(500) NOT NULL CONSTRAINT DF_TabEtiFormatiLayout_eti_t001 DEFAULT (''); " +
                       "IF COL_LENGTH('TabEtiFormatiLayout', 'eti_t002') IS NULL " +
                       "ALTER TABLE TabEtiFormatiLayout ADD eti_t002 nvarchar(500) NOT NULL CONSTRAINT DF_TabEtiFormatiLayout_eti_t002 DEFAULT (''); " +
                       "IF COL_LENGTH('TabEtiFormatiLayout', 'eti_t003') IS NULL " +
                       "ALTER TABLE TabEtiFormatiLayout ADD eti_t003 nvarchar(500) NOT NULL CONSTRAINT DF_TabEtiFormatiLayout_eti_t003 DEFAULT (''); " +
                       "IF COL_LENGTH('TabEtiFormatiLayout', 'eti_t004') IS NULL " +
                       "ALTER TABLE TabEtiFormatiLayout ADD eti_t004 nvarchar(500) NOT NULL CONSTRAINT DF_TabEtiFormatiLayout_eti_t004 DEFAULT (''); " +
                       "IF COL_LENGTH('TabEtiFormatiLayout', 'eti_t005') IS NULL " +
                       "ALTER TABLE TabEtiFormatiLayout ADD eti_t005 nvarchar(500) NOT NULL CONSTRAINT DF_TabEtiFormatiLayout_eti_t005 DEFAULT (''); " +
                       "IF COL_LENGTH('TabEtiFormatiLayout', 'eti_t006') IS NULL " +
                       "ALTER TABLE TabEtiFormatiLayout ADD eti_t006 nvarchar(500) NOT NULL CONSTRAINT DF_TabEtiFormatiLayout_eti_t006 DEFAULT (''); " +
                       "IF COL_LENGTH('TabEtiFormatiLayout', 'eti_t007') IS NULL " +
                       "ALTER TABLE TabEtiFormatiLayout ADD eti_t007 nvarchar(500) NOT NULL CONSTRAINT DF_TabEtiFormatiLayout_eti_t007 DEFAULT (''); " +
                       "IF COL_LENGTH('TabEtiFormatiLayout', 'eti_t008') IS NULL " +
                       "ALTER TABLE TabEtiFormatiLayout ADD eti_t008 nvarchar(500) NOT NULL CONSTRAINT DF_TabEtiFormatiLayout_eti_t008 DEFAULT (''); " +
                       "IF COL_LENGTH('TabEtiFormatiLayout', 'eti_h001') IS NULL " +
                       "ALTER TABLE TabEtiFormatiLayout ADD eti_h001 nvarchar(200) NOT NULL CONSTRAINT DF_TabEtiFormatiLayout_eti_h001 DEFAULT (''); " +
                       "IF COL_LENGTH('TabEtiFormatiLayout', 'eti_h002') IS NULL " +
                       "ALTER TABLE TabEtiFormatiLayout ADD eti_h002 nvarchar(200) NOT NULL CONSTRAINT DF_TabEtiFormatiLayout_eti_h002 DEFAULT (''); " +
                       "IF COL_LENGTH('TabEtiFormatiLayout', 'eti_h003') IS NULL " +
                       "ALTER TABLE TabEtiFormatiLayout ADD eti_h003 nvarchar(200) NOT NULL CONSTRAINT DF_TabEtiFormatiLayout_eti_h003 DEFAULT (''); " +
                       "IF COL_LENGTH('TabEtiFormatiLayout', 'eti_h004') IS NULL " +
                       "ALTER TABLE TabEtiFormatiLayout ADD eti_h004 nvarchar(200) NOT NULL CONSTRAINT DF_TabEtiFormatiLayout_eti_h004 DEFAULT (''); " +
                       "IF COL_LENGTH('TabEtiFormatiLayout', 'eti_h005') IS NULL " +
                       "ALTER TABLE TabEtiFormatiLayout ADD eti_h005 nvarchar(200) NOT NULL CONSTRAINT DF_TabEtiFormatiLayout_eti_h005 DEFAULT (''); " +
                       "IF COL_LENGTH('TabEtiFormatiLayout', 'eti_h006') IS NULL " +
                       "ALTER TABLE TabEtiFormatiLayout ADD eti_h006 nvarchar(200) NOT NULL CONSTRAINT DF_TabEtiFormatiLayout_eti_h006 DEFAULT (''); " +
                       "IF COL_LENGTH('TabEtiFormatiLayout', 'eti_h007') IS NULL " +
                       "ALTER TABLE TabEtiFormatiLayout ADD eti_h007 nvarchar(200) NOT NULL CONSTRAINT DF_TabEtiFormatiLayout_eti_h007 DEFAULT (''); " +
                       "IF COL_LENGTH('TabEtiFormatiLayout', 'eti_h008') IS NULL " +
                       "ALTER TABLE TabEtiFormatiLayout ADD eti_h008 nvarchar(200) NOT NULL CONSTRAINT DF_TabEtiFormatiLayout_eti_h008 DEFAULT (''); " +
                       "END";

            _clsFun.SqlWrite(s, strConSql);
        }

        private void MigraV002_TabEtiFormatiDefault(string strConSql)
        {
            // Elimina e reinserisce se esistono ancora con dati vuoti/incompleti
            // (verifica tramite s010 ='prezzo vecchio' per sapere se è già il formato professionale)
            string sCk = "SELECT COUNT(*) FROM TabEtiFormatiLayout WHERE eti_cod='900' AND eti_s010 != ''";
            DataTable tCk = _clsFun.FillTabSql("chk", sCk, true, strConSql);
            bool giaProfessionale = (tCk.Rows.Count > 0 && Convert.ToInt32(tCk.Rows[0][0]) > 0);

            string sCount = "SELECT COUNT(*) FROM TabEtiFormatiLayout";
            DataTable t = _clsFun.FillTabSql("check", sCount, true, strConSql);
            bool vuoto = (t.Rows.Count > 0 && Convert.ToInt32(t.Rows[0][0]) == 0);

            if (vuoto || !giaProfessionale)
            {
                // Rimuovi i vecchi esempi incompleti se presenti
                _clsFun.SqlWrite("DELETE FROM TabEtiFormatiLayout WHERE eti_cod IN ('900','901')", strConSql);

                // === FORMATO 900 — Etichetta da Scaffale 100x50mm (singola, stampa diretta) ===
                // Formato stringa: "Visibile(S/N);Y,X,Larghezza,Altezza,Align;FontNome|Dimensione|Stile"
                string s003 = "S;1,2,96,7,0;Arial|6|N";           // Intestazione/RagSoc
                string s004 = "S;9,2,96,11,0;Arial|9|B";          // Descrizione Articolo (grassetto)
                string s007 = "S;20,2,60,6,0;Arial|6|N";          // Slogan/Data Promo
                string s008 = "S;20,64,34,6,1;Arial|6|N";         // Prezzo al Kg
                string s010 = "S;27,2,28,7,0;Arial|7|N";          // Prezzo VECCHIO barrato
                string s011 = "S;26,32,32,12,0;Arial|14|B";       // PREZZO FINALE grande
                string s013 = "S;38,2,55,11,0;Arial|7|N";         // Barcode EAN
                string s014 = "S;15,64,34,5,1;Arial|5|N";         // Origine
                string s015 = "S;18,64,34,5,1;Arial|5|N";         // Calibro
                string s016 = "S;38,64,34,5,1;Arial|6|B";         // PLU

                string sq = "INSERT INTO TabEtiFormatiLayout " +
                    "(eti_cod,eti_des,eti_tip,eti_dim,eti_pag,eti_s003,eti_s004,eti_s005,eti_s006,eti_s007,eti_s008,eti_s009,eti_s010,eti_s011,eti_s012,eti_s013,eti_s014,eti_s015,eti_s016,eti_ann) " +
                    "VALUES ('900','Scaffale 100x50 mm','PDF','100,50','1,1,0,0,100,50'," +
                    "'" + s003 + "','" + s004 + "','','','" + s007 + "','" + s008 + "','','" + s010 + "','" + s011 + "','','" + s013 + "','" + s014 + "','" + s015 + "','" + s016 + "',0)";
                _clsFun.SqlWrite(sq, strConSql);

                // === FORMATO 901 — Foglio A4 colonne multiple (70x35mm per etichetta, 3 col x 8 rig) ===
                string s003b = "S;1,2,66,5,0;Arial|5|N";
                string s004b = "S;6,2,66,8,0;Arial|7|B";
                string s010b = "S;15,2,20,5,0;Arial|6|N";
                string s011b = "S;14,24,22,10,0;Arial|12|B";
                string s013b = "S;26,2,40,8,0;Arial|6|N";
                string s014b = "S;26,45,22,4,1;Arial|4|N"; // Origine piccola

                string sqb = "INSERT INTO TabEtiFormatiLayout " +
                    "(eti_cod,eti_des,eti_tip,eti_dim,eti_pag,eti_s003,eti_s004,eti_s005,eti_s006,eti_s007,eti_s008,eti_s009,eti_s010,eti_s011,eti_s012,eti_s013,eti_s014,eti_s015,eti_s016,eti_ann) " +
                    "VALUES ('901','A4 (3 col x 8 rig)','PDF','210,297','3,8,2,2,70,35'," +
                    "'" + s003b + "','" + s004b + "','','','','','','','" + s010b + "','" + s011b + "','','" + s013b + "','" + s014b + "','','',0)";
                _clsFun.SqlWrite(sqb, strConSql);
            }
        }
        // -------------------------------------------------------------------
        // V003 — Allarga AnaArticoli.art_img a nvarchar(255)
        //        (colonna originale era nvarchar(50), troppo corta per percorsi)
        // -------------------------------------------------------------------
        private void MigraV003_AnaArticoli_art_img(string strConSql)
        {
            // Controlla la lunghezza attuale: COL_LENGTH restituisce byte, nvarchar => 2 byte/char
            // Se è già >= 400 (200 chars * 2) non tocca nulla
            string s = "IF EXISTS (" +
                       "  SELECT 1 FROM sys.columns c " +
                       "  JOIN sys.objects o ON c.object_id = o.object_id " +
                       "  WHERE o.name = 'AnaArticoli' AND c.name = 'art_img' AND c.max_length < 400" +
                       ") " +
                       "ALTER TABLE [dbo].[AnaArticoli] ALTER COLUMN [art_img] nvarchar(255) NOT NULL";
            _clsFun.SqlWrite(s, strConSql);
        }

        // -------------------------------------------------------------------
        // V004 — Allarga TabEtiFormatiLayout.eti_s012 e eti_s017 a nvarchar(1000)
        //        (permettere percorsi immagine molto lunghi nei formati etichette)
        // -------------------------------------------------------------------
        private void MigraV004_TabEtiFormatiLayout_ImgLengths(string strConSql)
        {
            string s1 = "IF EXISTS (" +
                        "  SELECT 1 FROM sys.columns c " +
                        "  JOIN sys.objects o ON c.object_id = o.object_id " +
                        "  WHERE o.name = 'TabEtiFormatiLayout' AND c.name = 'eti_s012' AND c.max_length < 2000" +
                        ") " +
                        "ALTER TABLE [dbo].[TabEtiFormatiLayout] ALTER COLUMN [eti_s012] nvarchar(1000) NOT NULL";
            _clsFun.SqlWrite(s1, strConSql);

            string s2 = "IF EXISTS (" +
                        "  SELECT 1 FROM sys.columns c " +
                        "  JOIN sys.objects o ON c.object_id = o.object_id " +
                        "  WHERE o.name = 'TabEtiFormatiLayout' AND c.name = 'eti_s017' AND c.max_length < 2000" +
                        ") " +
                        "ALTER TABLE [dbo].[TabEtiFormatiLayout] ALTER COLUMN [eti_s017] nvarchar(1000) NOT NULL";
            _clsFun.SqlWrite(s2, strConSql);
        }

        private void MigraV005_TabParametri_Completa(string strConSql)
        {
            string[,] defs = new string[,]
            {
                { "001", "Codice Fornitore di default", "", "0" },
                { "002", "Gestione download fornitori", "", "0" },
                { "003", "Tempo di ritardo chiusure per lettura venduto 3000", "500000000", "0" },
                { "004", "Divisione ordine fornitori", "1,2", "0" },
                { "005", "Modalità grafica", "S", "0" },
                { "006", "PREZZI A ZERO", "", "0" },
                { "007", "PREFISSO COD BILANCIA", "", "0" },
                { "008", "OMEGA LUNG CODICE", "3", "0" },
                { "009", "Statistiche pro", "N", "0" },
                { "010", "ECR di defalut", "001,001,001", "0" },
                { "011", "Codice articolo modificabile in inserimento", "N", "0" },
                { "012", "Codice IVA di default", "022", "0" },
                { "013", "Codice reparto di default", "099", "0" },
                { "014", "Aggiornamento prezzi solo da fattura elettronica", "", "0" },
                { "015", "Agg des div for primario", "S", "0" },
                { "016", "PERCORSO DIVULGAZIONE NEGOZI", "N", "0" },
                { "017", "Controllo ECR da divulgazione fornitori", "", "0" },
                { "018", "Disattivazione inserimento articolo", "", "0" },
                { "019", "Variazione forzata da div fornitori", "", "0" },
                { "020", "Esclusione barcode pref 20 21 da div", "", "0" },
                { "021", "Tipo gestione ingredienti bilancia", "", "0" },
                { "022", "Stampa documenti con raggruppamento articoli", "", "0" },
                { "023", "Divulgazione clienti alle casse per fattura", "", "0" },
                { "024", "Decimali su peso e importo in stampa DDT", "", "0" },
                { "025", "Doppio documento su stessa pagina", "", "0" },
                { "026", "Parametri inserimento fattura da cassa", "", "0" },
                { "027", "Tipo fidelity multitessera", "", "0" },
                { "028", "Parametri bilance prezzo modificabile", "", "0" },
                { "029", "Parametri percorso file WEB art_web", "", "0" },
                { "030", "Parametri percorso file ingredienti", "", "0" },
                { "031", "Gestione lotti con divulgazione alle casse", "", "0" },
                { "032", "Regola import articoli eanizzati", "", "0" },
                { "033", "Fatt elettronica", "S", "0" },
                { "034", "Parametri fatture differite", "", "0" },
                { "035", "Parametri lettura scontrino bilancia", "", "0" },
                { "036", "Divulgazione a sede remota", "", "0" },
                { "037", "Parametri gestione celiaci", "", "0" },
                { "038", "Parametri e gestione password operatore", "", "0" },
                { "039", "Divulgazioni in cassa solo prezzi negozio", "", "0" }
            };
            string sColCheck = "IF COL_LENGTH('TabParametri', 'tab_ann') IS NULL ALTER TABLE TabParametri ADD tab_ann bit NOT NULL CONSTRAINT DF_TabParametri_tab_ann DEFAULT ((0))";
            _clsFun.SqlWrite(sColCheck, strConSql);

            string sSql = "SELECT tab_cod FROM TabParametri";
            DataTable tCurr = _clsFun.FillTabSql("TabParametri", sSql, false, strConSql);

            System.Collections.Generic.HashSet<string> existingCodes = new System.Collections.Generic.HashSet<string>();
            if (tCurr != null)
            {
                foreach (DataRow r in tCurr.Rows)
                {
                    if (r["tab_cod"] != DBNull.Value)
                    {
                        existingCodes.Add(r["tab_cod"].ToString().Trim().PadLeft(3, '0'));
                    }
                }
            }

            for (int i = 1; i <= 999; i++)
            {
                string codStr = i.ToString("000");
                if (!existingCodes.Contains(codStr))
                {
                    string des = "Parametro " + codStr + " non utilizzato";
                    string val = "";
                    string ann = "1";

                    for (int k = 0; k < defs.GetLength(0); k++)
                    {
                        if (defs[k, 0] == codStr)
                        {
                            des = defs[k, 1];
                            val = defs[k, 2];
                            ann = defs[k, 3];
                            break;
                        }
                    }

                    string sq = "INSERT INTO TabParametri (tab_cod, tab_des, tab_val, tab_ann) VALUES ('" + codStr + "', '" + des.Replace("'", "''") + "', '" + val.Replace("'", "''") + "', " + ann + ")";
                    _clsFun.SqlWrite(sq, strConSql);
                }
            }
        }

        private void MigraV006_Indexes_Optimization(string strConSql)
        {
            try
            {
                string s1 = "IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = 'IX_AnaBarcode_ean_art' AND object_id = OBJECT_ID('AnaBarcode')) " +
                            "CREATE NONCLUSTERED INDEX IX_AnaBarcode_ean_art ON AnaBarcode(ean_art) INCLUDE (ean_ean, ean_prv, ean_qta);";
                _clsFun.SqlWrite(s1, strConSql);

                string s2 = "IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = 'IX_GesLisAcquisto_lia_art' AND object_id = OBJECT_ID('GesLisAcquisto')) " +
                            "CREATE NONCLUSTERED INDEX IX_GesLisAcquisto_lia_art ON GesLisAcquisto(lia_art, lia_ann) INCLUDE (lia_for, lia_cos, lia_arf, lia_pxc);";
                _clsFun.SqlWrite(s2, strConSql);

                string s3 = "IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = 'IX_GesLisVendita_liv_art' AND object_id = OBJECT_ID('GesLisVendita')) " +
                            "CREATE NONCLUSTERED INDEX IX_GesLisVendita_liv_art ON GesLisVendita(liv_art, liv_lis, liv_ann) INCLUDE (liv_prv, liv_sta);";
                _clsFun.SqlWrite(s3, strConSql);

                string s4 = "IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = 'IX_AnaArticoli_art_des' AND object_id = OBJECT_ID('AnaArticoli')) " +
                            "CREATE NONCLUSTERED INDEX IX_AnaArticoli_art_des ON AnaArticoli(art_des) INCLUDE (art_cod, art_sta, art_rep);";
                _clsFun.SqlWrite(s4, strConSql);

                string s5 = "IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = 'IX_AnaArticoli_art_rep' AND object_id = OBJECT_ID('AnaArticoli')) " +
                            "CREATE NONCLUSTERED INDEX IX_AnaArticoli_art_rep ON AnaArticoli(art_rep);";
                _clsFun.SqlWrite(s5, strConSql);

                string s6 = "IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = 'IX_AnaArticoli_art_ec1' AND object_id = OBJECT_ID('AnaArticoli')) " +
                            "CREATE NONCLUSTERED INDEX IX_AnaArticoli_art_ec1 ON AnaArticoli(art_ec1);";
                _clsFun.SqlWrite(s6, strConSql);

                string s7 = "IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = 'IX_GesVariazioni_var_art' AND object_id = OBJECT_ID('GesVariazioni')) " +
                            "CREATE NONCLUSTERED INDEX IX_GesVariazioni_var_art ON GesVariazioni(var_art, var_tip, var_num);";
                _clsFun.SqlWrite(s7, strConSql);

                string s8 = "IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = 'IX_LogLog_log_dti' AND object_id = OBJECT_ID('LogLog')) " +
                            "CREATE NONCLUSTERED INDEX IX_LogLog_log_dti ON LogLog(log_dti);";
                _clsFun.SqlWrite(s8, strConSql);
            }
            catch (Exception ex)
            {
                _clsFun.ErrorLog(ex.Message, "MigraV006_Indexes_Optimization");
            }
        }

        private void MigraV007_TabEti_StatoPostStampa(string strConSql)
        {
            try
            {
                // 1. TabEtiFormatiLayout: aggiunta colonna eti_sta (S=Segnala/chiedi stampata, N=da stampare, E=eliminata)
                string sql1 = "IF EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[TabEtiFormatiLayout]') AND type in (N'U')) " +
                              "BEGIN " +
                              "  IF COL_LENGTH('TabEtiFormatiLayout', 'eti_sta') IS NULL " +
                              "    ALTER TABLE TabEtiFormatiLayout ADD eti_sta nvarchar(1) NOT NULL CONSTRAINT DF_TabEtiFormatiLayout_eti_sta DEFAULT ('S'); " +
                              "END";
                _clsFun.SqlWrite(sql1, strConSql);

                // 2. TabEtichette: aggiunta colonna tab_sta (S=Segnala/chiedi stampata, N=da stampare, E=eliminata)
                string sql2 = "IF EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[TabEtichette]') AND type in (N'U')) " +
                              "BEGIN " +
                              "  IF COL_LENGTH('TabEtichette', 'tab_sta') IS NULL " +
                              "    ALTER TABLE TabEtichette ADD tab_sta nvarchar(1) NOT NULL CONSTRAINT DF_TabEtichette_tab_sta DEFAULT ('S'); " +
                              "END";
                _clsFun.SqlWrite(sql2, strConSql);
            }
            catch (Exception ex)
            {
                _clsFun.ErrorLog(ex.Message, "MigraV007_TabEti_StatoPostStampa");
            }
        }

        private void MigraV008_TabEti_TestoLibero(string strConSql)
        {
            try
            {
                string s = "IF EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[TabEtiFormatiLayout]') AND type in (N'U')) " +
                           "BEGIN " +
                           "IF COL_LENGTH('TabEtiFormatiLayout', 'eti_t001') IS NULL ALTER TABLE TabEtiFormatiLayout ADD eti_t001 nvarchar(500) NOT NULL CONSTRAINT DF_TabEtiFormatiLayout_eti_t001 DEFAULT (''); " +
                           "IF COL_LENGTH('TabEtiFormatiLayout', 'eti_t002') IS NULL ALTER TABLE TabEtiFormatiLayout ADD eti_t002 nvarchar(500) NOT NULL CONSTRAINT DF_TabEtiFormatiLayout_eti_t002 DEFAULT (''); " +
                           "IF COL_LENGTH('TabEtiFormatiLayout', 'eti_t003') IS NULL ALTER TABLE TabEtiFormatiLayout ADD eti_t003 nvarchar(500) NOT NULL CONSTRAINT DF_TabEtiFormatiLayout_eti_t003 DEFAULT (''); " +
                           "IF COL_LENGTH('TabEtiFormatiLayout', 'eti_t004') IS NULL ALTER TABLE TabEtiFormatiLayout ADD eti_t004 nvarchar(500) NOT NULL CONSTRAINT DF_TabEtiFormatiLayout_eti_t004 DEFAULT (''); " +
                           "IF COL_LENGTH('TabEtiFormatiLayout', 'eti_t005') IS NULL ALTER TABLE TabEtiFormatiLayout ADD eti_t005 nvarchar(500) NOT NULL CONSTRAINT DF_TabEtiFormatiLayout_eti_t005 DEFAULT (''); " +
                           "IF COL_LENGTH('TabEtiFormatiLayout', 'eti_t006') IS NULL ALTER TABLE TabEtiFormatiLayout ADD eti_t006 nvarchar(500) NOT NULL CONSTRAINT DF_TabEtiFormatiLayout_eti_t006 DEFAULT (''); " +
                           "IF COL_LENGTH('TabEtiFormatiLayout', 'eti_t007') IS NULL ALTER TABLE TabEtiFormatiLayout ADD eti_t007 nvarchar(500) NOT NULL CONSTRAINT DF_TabEtiFormatiLayout_eti_t007 DEFAULT (''); " +
                           "IF COL_LENGTH('TabEtiFormatiLayout', 'eti_t008') IS NULL ALTER TABLE TabEtiFormatiLayout ADD eti_t008 nvarchar(500) NOT NULL CONSTRAINT DF_TabEtiFormatiLayout_eti_t008 DEFAULT (''); " +
                           "END";
                _clsFun.SqlWrite(s, strConSql);
            }
            catch (Exception ex)
            {
                _clsFun.ErrorLog(ex.Message, "MigraV008_TabEti_TestoLibero");
            }
        }

        private void MigraV009_CleanPlaceholderAndScaleOptimization(string strConSql)
        {
            try
            {
                // 1. Pulizia approfondita record fittizi/placeholder lasciati da import errati ('NUOVO', '[NUOVO]', 'NEW', '[NEW]', 'AUTO', '[AUTO]', 'N', '[N]')
                string sCleanEan = "DELETE FROM AnaBarcode WHERE UPPER(LTRIM(RTRIM(ean_art))) IN ('NUOVO', '[NUOVO]', 'NEW', '[NEW]', 'AUTO', '[AUTO]', 'N', '[N]') OR UPPER(LTRIM(RTRIM(ean_art))) LIKE '%NUOVO%' OR UPPER(LTRIM(RTRIM(ean_ean))) IN ('NUOVO', '[NUOVO]', 'NEW', '[NEW]', 'AUTO', '[AUTO]', 'N', '[N]');";
                _clsFun.SqlWrite(sCleanEan, strConSql);

                string sCleanLia = "DELETE FROM GesLisAcquisto WHERE UPPER(LTRIM(RTRIM(lia_art))) IN ('NUOVO', '[NUOVO]', 'NEW', '[NEW]', 'AUTO', '[AUTO]', 'N', '[N]') OR UPPER(LTRIM(RTRIM(lia_art))) LIKE '%NUOVO%';";
                _clsFun.SqlWrite(sCleanLia, strConSql);

                string sCleanLiv = "DELETE FROM GesLisVendita WHERE UPPER(LTRIM(RTRIM(liv_art))) IN ('NUOVO', '[NUOVO]', 'NEW', '[NEW]', 'AUTO', '[AUTO]', 'N', '[N]') OR UPPER(LTRIM(RTRIM(liv_art))) LIKE '%NUOVO%';";
                _clsFun.SqlWrite(sCleanLiv, strConSql);

                string sCleanVar = "DELETE FROM GesVariazioni WHERE UPPER(LTRIM(RTRIM(var_art))) IN ('NUOVO', '[NUOVO]', 'NEW', '[NEW]', 'AUTO', '[AUTO]', 'N', '[N]') OR UPPER(LTRIM(RTRIM(var_art))) LIKE '%NUOVO%';";
                _clsFun.SqlWrite(sCleanVar, strConSql);

                string sCleanArt = "DELETE FROM AnaArticoli WHERE UPPER(LTRIM(RTRIM(art_cod))) IN ('NUOVO', '[NUOVO]', 'NEW', '[NEW]', 'AUTO', '[AUTO]', 'N', '[N]') OR UPPER(LTRIM(RTRIM(art_cod))) LIKE '%NUOVO%';";
                _clsFun.SqlWrite(sCleanArt, strConSql);

                // 2. Pulizia barcode orfani e codici vuoti
                string sCleanOrphans = @"
                    DELETE FROM AnaBarcode WHERE (ean_art IS NULL OR LTRIM(RTRIM(ean_art)) = '') OR (ean_art NOT IN (SELECT art_cod FROM AnaArticoli));
                    DELETE FROM GesLisAcquisto WHERE (lia_art IS NULL OR LTRIM(RTRIM(lia_art)) = '') OR (lia_art NOT IN (SELECT art_cod FROM AnaArticoli));
                    DELETE FROM GesLisVendita WHERE (liv_art IS NULL OR LTRIM(RTRIM(liv_art)) = '') OR (liv_art NOT IN (SELECT art_cod FROM AnaArticoli));
                    DELETE FROM GesVariazioni WHERE (var_art IS NULL OR LTRIM(RTRIM(var_art)) = '') OR (var_art NOT IN (SELECT art_cod FROM AnaArticoli));
                    DELETE FROM AnaArticoli WHERE art_cod IS NULL OR LTRIM(RTRIM(art_cod)) = '';
                ";
                _clsFun.SqlWrite(sCleanOrphans, strConSql);

                // 3. Sanitizzazione flag ean_bil in AnaBarcode (solo prefissi bilancia '2' possono avere ean_bil=1)
                string sFixEanBil = "UPDATE AnaBarcode SET ean_bil = 0 WHERE ean_bil = 1 AND ean_ean NOT LIKE '2%';";
                _clsFun.SqlWrite(sFixEanBil, strConSql);

                // 4. Allineamento flag art_bil e default obbligatori su AnaArticoli
                string sFixArtFields = @"
                    UPDATE AnaArticoli SET art_bil = 1 WHERE (art_bil = 0 OR art_bil IS NULL) AND ((art_plu IS NOT NULL AND LTRIM(RTRIM(art_plu)) != '' AND LTRIM(RTRIM(art_plu)) != '0') OR (art_tas IS NOT NULL AND LTRIM(RTRIM(art_tas)) != '' AND LTRIM(RTRIM(art_tas)) != '0'));
                    UPDATE AnaArticoli SET art_deb = LEFT(art_des, 20) WHERE (art_deb IS NULL OR LTRIM(RTRIM(art_deb)) = '') AND art_des IS NOT NULL AND LTRIM(RTRIM(art_des)) != '';
                    UPDATE AnaArticoli SET art_umi = 'NR' WHERE art_umi IS NULL OR LTRIM(RTRIM(art_umi)) = '';
                    UPDATE AnaArticoli SET art_umc = 'NR' WHERE art_umc IS NULL OR LTRIM(RTRIM(art_umc)) = '';
                    UPDATE AnaArticoli SET art_tgr = 'PZ' WHERE art_tgr IS NULL OR LTRIM(RTRIM(art_tgr)) = '';
                    UPDATE AnaArticoli SET art_sta = 'A' WHERE art_sta IS NULL OR LTRIM(RTRIM(art_sta)) = '';
                ";
                _clsFun.SqlWrite(sFixArtFields, strConSql);

                // 5. Riallineamento Numeratore Automatico Articoli in TabNumeratori
                string sFixNum = @"
                    DECLARE @maxCod int;
                    SELECT @maxCod = MAX(CAST(art_cod AS int)) FROM AnaArticoli WHERE ISNUMERIC(art_cod) = 1 AND art_cod LIKE '0%' AND LEN(RTRIM(art_cod)) = 7;
                    IF @maxCod IS NOT NULL AND @maxCod > 0
                    BEGIN
                        DECLARE @currVal int;
                        SELECT @currVal = CAST(tab_val AS int) FROM TabNumeratori WHERE tab_yea = '0000' AND tab_cod = '001' AND ISNUMERIC(tab_val) = 1;
                        IF @currVal IS NULL OR @maxCod > @currVal
                        BEGIN
                            IF EXISTS (SELECT 1 FROM TabNumeratori WHERE tab_yea = '0000' AND tab_cod = '001')
                                UPDATE TabNumeratori SET tab_val = CAST(@maxCod AS varchar(20)) WHERE tab_yea = '0000' AND tab_cod = '001';
                            ELSE
                                INSERT INTO TabNumeratori (tab_yea, tab_cod, tab_des, tab_val, tab_ann) VALUES ('0000', '001', 'NUMERATORE ARTICOLI', CAST(@maxCod AS varchar(20)), 0);
                        END
                    END
                ";
                _clsFun.SqlWrite(sFixNum, strConSql);

                // 6. Indici per ricerca rapida Tasto Bilancia e Codice PLU
                string sIdxPlu = "IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = 'IX_AnaArticoli_art_plu' AND object_id = OBJECT_ID('AnaArticoli')) " +
                                 "CREATE NONCLUSTERED INDEX IX_AnaArticoli_art_plu ON AnaArticoli(art_plu, art_reb) INCLUDE (art_cod, art_des, art_sta, art_bil);";
                _clsFun.SqlWrite(sIdxPlu, strConSql);

                string sIdxTas = "IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = 'IX_AnaArticoli_art_tas' AND object_id = OBJECT_ID('AnaArticoli')) " +
                                 "CREATE NONCLUSTERED INDEX IX_AnaArticoli_art_tas ON AnaArticoli(art_tas) INCLUDE (art_cod, art_des, art_sta, art_bil);";
                _clsFun.SqlWrite(sIdxTas, strConSql);

                string sIdxEanBil = "IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = 'IX_AnaBarcode_ean_bil' AND object_id = OBJECT_ID('AnaBarcode')) " +
                                    "CREATE NONCLUSTERED INDEX IX_AnaBarcode_ean_bil ON AnaBarcode(ean_bil, ean_art) INCLUDE (ean_ean, ean_prv);";
                _clsFun.SqlWrite(sIdxEanBil, strConSql);
            }
            catch (Exception ex)
            {
                _clsFun.ErrorLog(ex.Message, "MigraV009_CleanPlaceholderAndScaleOptimization");
            }
        }
    }
}
