using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;

namespace APOffice
{
    public partial class frmGesFattElettroniche : Form
    {
        clsDefine _clsDef = new clsDefine();
        clsFuncs _clsFun = new clsFuncs();
        clsQuery _clsQry = new clsQuery();

        private string XmlEsc(string value)
        {
            if (string.IsNullOrEmpty(value)) return "";
            return System.Security.SecurityElement.Escape(value);
        }

        private string _strPath = "C:\\ApProject\\Temp\\DivDocumenti\\";

        private string _strConSql = "";

        public string _strYea = "";
        public string _strSuf = "";

        public frmGesFattElettroniche()
        {
            InitializeComponent();
            new clsGesGraph().SetGraph(this, 0);
            _strConSql = _clsFun.ConSql("");
            this.FormClosing += new FormClosingEventHandler(this.frmGesFattElettroniche_FormClosing);
        }

        private void frmGesFattElettroniche_FormClosing(object sender, FormClosingEventArgs e)
        {
            clsUiIcons.SaveFormBounds(this);
            clsUiIcons.SaveGridColumnWidths(dgv1, "frmGesFattElettroniche_dgv1");
        }

        private void frmGesFattElettroniche_Load(object sender, EventArgs e)
        {
            clsUiIcons.RestoreFormBounds(this);
            SetDgv1();
            clsUiIcons.RestoreGridColumnWidths(dgv1, "frmGesFattElettroniche_dgv1");
            FillTab();
            FillAnag();
            FillDati("DAINV");
            ApplyModernUi();
        }

        private void ApplyModernUi()
        {
            try
            {
                if (menuStrip1 != null)
                {
                    menuStrip1.Renderer = clsUiIcons.GetModernMenuRenderer();

                    if (esciToolStripMenuItem != null)
                        esciToolStripMenuItem.Image = clsUiIcons.GetIcon("exit", 16);
                }

                if (dgv1 != null)
                    clsUiIcons.StyleDataGridView(dgv1);

                if (btnOk != null)
                {
                    clsUiIcons.StyleButton(btnOk, clsUiIcons.GetIcon("xml", 22), Color.FromArgb(30, 64, 175), Color.White);
                    btnOk.Font = new Font("Segoe UI", 10.5f, FontStyle.Bold);
                }
            }
            catch (Exception ex)
            {
                _clsFun.ErrorLog("ApplyModernUi", ex.Message);
            }
        }

        private void esciToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Esci();
        }

        private void frmGesFattElettroniche_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Escape)
                Esci();
        }

        private void Esci()
        {
            this.Close();
        }

        private void SetDgv1()
        {
            DataTable tSta = _clsFun.FillTabSql("TabStatoDivulgazioni", "SELECT * FROM TabStatoDivulgazioni WHERE tab_var=1", false, _strConSql);

            dgv1.AutoGenerateColumns = false;
            //dgv1.VirtualMode = true;
            //dgv1.Dock = DockStyle.Fill;
            dgv1.AllowUserToAddRows = false;
            dgv1.ReadOnly = false;
            dgv1.AllowUserToDeleteRows = false;
            dgv1.AllowUserToResizeColumns = true;
            //dgv1.DisplayedRowCount() = true;

            DataGridViewTextBoxColumn cTbc;
            //DataGridViewCheckBoxColumn cCbc;
            //DataGridViewComboBoxColumn cCmb;

            //cTbc = new DataGridViewTextBoxColumn();
            //cTbc.DataPropertyName = "mov_idx";
            //cTbc.Name = "Idx";
            //cTbc.Width = 20;
            //cTbc.ValueType = typeof(string);
            //cTbc.ReadOnly = true;
            ////cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            //dgv1.Columns.Add(cTbc);

            //cCmb = new DataGridViewComboBoxColumn();
            //cCmb.DataPropertyName = "DocInv";
            //cCmb.Name = "Stato";
            //cCmb.Width = 60;
            //cCmb.DataSource = tSta;
            //cCmb.ValueMember = "tab_cod";
            //cCmb.DisplayMember = "tab_des";
            //cCmb.ReadOnly = true;
            //cCmb.DisplayStyle = DataGridViewComboBoxDisplayStyle.Nothing;
            //cCmb.DefaultCellStyle.Font = new Font("Microsoft Sans", 8.75F, GraphicsUnit.Pixel);
            //dgv1.Columns.Add((DataGridViewColumn)cCmb);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "FatSel";
            cTbc.Name = "Sel";
            cTbc.Width = 30;
            cTbc.ValueType = typeof(string);
            cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            cTbc.ReadOnly = true;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "fat_tdo";
            cTbc.Name = "Tipo documento";
            cTbc.Width = 50;
            cTbc.ValueType = typeof(string);
            cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            cTbc.ReadOnly = true;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "fat_ndo";
            cTbc.Name = "N.Documento";
            cTbc.Width = 80;
            cTbc.ValueType = typeof(string);
            cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            cTbc.ReadOnly = true;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "fat_ddo";
            cTbc.Name = "D.Documento";
            cTbc.Width = 80;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = true;
            cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "fat_cfo";
            cTbc.Name = "Cliente";
            cTbc.Width = 60;
            cTbc.ValueType = typeof(string);
            cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            cTbc.ReadOnly = true;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "CliRag";
            cTbc.Name = "Descrizione";
            cTbc.Width = 250;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = true;
            //cTbc.MaxInputLength = 50;
            //cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dgv1.Columns.Add(cTbc);

            //cTbc = new DataGridViewTextBoxColumn();
            //cTbc.DataPropertyName = "FatTpd";
            //cTbc.Name = "Doc. fatt. elettronica";
            //cTbc.Width = 60;
            //cTbc.ValueType = typeof(string);
            //cTbc.ReadOnly = true;
            ////cTbc.MaxInputLength = 50;
            ////cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            //cTbc.ToolTipText = "Numero identificativo fattura elettronica";
            //cTbc.DefaultCellStyle.ForeColor = Color.Blue;
            //dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "fat_ele";
            cTbc.Name = "N. identificativo";
            cTbc.Width = 60;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = true;
            //cTbc.MaxInputLength = 50;
            //cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            cTbc.ToolTipText = "Numero identificativo fattura elettronica";
            cTbc.DefaultCellStyle.ForeColor = Color.Blue;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "FatMsg";
            cTbc.Name = "Messaggi";
            cTbc.Width = 200;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = true;
            //cTbc.MaxInputLength = 50;
            //cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            cTbc.DefaultCellStyle.ForeColor = Color.Red;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "FatPag";
            cTbc.Name = "Tipo pagamento";
            cTbc.Width = 200;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = true;
            //cTbc.MaxInputLength = 50;
            //cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            //cTbc.DefaultCellStyle.ForeColor = Color.Red;
            dgv1.Columns.Add(cTbc);
        }

        private void FillTab()
        {
            for (int i = DateTime.Today.Year - 4; i < DateTime.Today.Year + 1; i++)
                cmbYea.Items.Add(i.ToString());
            cmbYea.SelectedIndex = 4;

            if(_strYea != "")
                cmbYea.Text = _strYea;
        }

        private void FillAnag()
        {
            string s = _clsFun.FileIni("R", clsDefine.enuIni.Ini09CodiceAzienda, "");
            if (s == "")
                s = "001";
            DataTable t = _clsQry.ConfAzienda(s);

            if(t.Rows.Count > 0)
            {
                lblCnfRag.Text = (string)t.Rows[0]["cnf_rag"];

                s = ((string)t.Rows[0]["cnf_ind"]).Replace("-", "") + " - ";
                s += ((string)t.Rows[0]["cnf_ncv"]).Replace("-", "") + " - ";
                s += ((string)t.Rows[0]["cnf_cap"]).Replace("-", "") + " - ";
                s += ((string)t.Rows[0]["cnf_loc"]).Replace("-", "") + " - ";
                s += ((string)t.Rows[0]["cnf_prv"]).Replace("-", "");
                lblCnfInd.Text =  s;

                lblCnfPiv.Text = (string)t.Rows[0]["cnf_piv"];
                lblCnfCfi.Text = (string)t.Rows[0]["cnf_cfi"];
                lblCnfMai.Text = (string)t.Rows[0]["cnf_pec"];
                lblCnfIba.Text = (string)t.Rows[0]["cnf_iba"];
                lblCnfBan.Text = (string)t.Rows[0]["cnf_ban"];
                lblCnfAgp.Text = (string)t.Rows[0]["cnf_agp"];
            }
        }

        private void FillDati(string strEle)
        {
            string s = "";
            string sWhe = " AND (fat_ele IS NULL OR fat_ele='')";
            if(strEle != "DAINV")
                sWhe = "";

            s = "SELECT * FROM TabDocTpd";
            DataTable tTpd = _clsFun.FillTabSql("TabDocTpd", s, false, _strConSql);

            // Ottimizzazione Fase 1: Pre-caricamento condizioni di pagamento in memoria
            s = "SELECT TabPagamenti.tab_cod, TabPagamenti.tab_des, TabPagamenti.tab_fte, TabPagamenti.tab_ftp, " +
                "TabPagamenti.tab_ann, TabPagamenti.tab_dpa, TabPagDate.tab_par AS DscPar " +
                "FROM TabPagamenti " +
                "LEFT OUTER JOIN TabPagDate ON TabPagamenti.tab_dpa = TabPagDate.tab_cod";
            DataTable tPagAll = _clsFun.FillTabSql("TabPagamentiAll", s, true, _strConSql);

            // Ottimizzazione Fase 1: Pre-caricamento verifica aliquote zero senza codice Natura
            s = "SELECT mov_yfa, mov_nfa, mov_iva FROM GesMovimenti " +
                "LEFT OUTER JOIN TabIva ON GesMovimenti.mov_iva = TabIva.tab_cod " +
                "WHERE tab_ali = 0 AND (tab_ele = '' OR tab_ele IS NULL) AND (tab_pos = '' OR tab_pos IS NULL) " +
                "AND mov_ann = 0 AND mov_yfa = '" + cmbYea.Text + "' AND mov_iva <> '' " +
                "GROUP BY mov_yfa, mov_nfa, mov_iva";
            DataTable tMovCheck = _clsFun.FillTabSql("GesMovCheck", s, false, _strConSql);
            Dictionary<string, string> dictIvaZeroSenzaNatura = new Dictionary<string, string>();
            if (tMovCheck != null)
            {
                foreach (DataRow rChk in tMovCheck.Rows)
                {
                    string key = Convert.ToString(rChk["mov_yfa"]) + "_" + Convert.ToString(rChk["mov_nfa"]);
                    if (!dictIvaZeroSenzaNatura.ContainsKey(key))
                        dictIvaZeroSenzaNatura[key] = Convert.ToString(rChk["mov_iva"]);
                }
            }

            s = "SELECT ";
            s += "fat_yfa, ";
            s += "fat_tdo, ";
            s += "fat_tpd, ";
            s += "fat_nfa, ";
            s += "fat_ndo, ";
            s += "fat_ddo, ";
            s += "fat_cfo, ";
            s += "fat_tpg, ";
            s += "fat_ele, ";
            s += "cli_des As CliRag, ";
            s += "cli_ind As CliInd, ";
            s += "cli_ncv As CliNcv, ";
            s += "cli_loc As CliLoc, ";
            s += "cli_cap As CliCap, ";
            s += "cli_prv As CliPrv, ";
            s += "cli_naz As CliNaz, ";
            s += "cli_pec As CliPec, ";
            s += "cli_sdi As CliSdi, ";
            s += "cli_ccp As CliCcp, ";
            s += "cli_piv As CliPiv, ";
            s += "cli_cfi As CliCfi, ";
            s += "cli_tip As CliTip ";
            s += "FROM GesFatTestate ";
            s += "LEFT JOIN AnaClienti ON GesFatTestate.fat_cfo = AnaClienti.cli_cod ";
            s += "WHERE fat_yfa='" + cmbYea.Text + "' AND fat_ann=0 AND fat_ndo <>'' AND (fat_tpd='FV' OR fat_tpd='FD') ";
            if(sWhe != "")
                s += sWhe;
            s += "ORDER BY fat_ndo";
            DataTable t = _clsFun.FillTabSql("GesFatTeste", s, false, _strConSql);

            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "FatSel",
                Caption = "Selezionato",
                MaxLength = 50,
                ReadOnly = false,
                DefaultValue = (String)""
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "FatMsg",
                Caption = "Messaggi",
                MaxLength = 150,
                ReadOnly = false,
                DefaultValue = (String)""
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "FatTpd",
                Caption = "Tipo documento",
                MaxLength = 4,
                ReadOnly = false,
                DefaultValue = (String)""
            });

            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "FatPag",
                Caption = "Tipi pagamento",
                MaxLength = 100,
                ReadOnly = false,
                DefaultValue = (String)""
            });

            foreach(DataRow y in t.Rows)
            {
                string sCliPiv = ((string)y["CliPiv"]).ToUpper();
                string sCliCfi = ((string)y["CliCfi"]).ToUpper();
                string sCliRag = ((string)y["CliRag"]).ToUpper();
                string sCliInd = ((string)y["CliInd"]).ToUpper();
                string sCliNcv = ((string)y["CliNcv"]).ToUpper();
                string sCliCap = ((string)y["CliCap"]).ToUpper();
                string sCliLoc = ((string)y["CliLoc"]).ToUpper();
                string sCliPrv = ((string)y["CliPrv"]).ToUpper();
                string sCliTip = Convert.ToString(y["CliTip"]);
                string sCliPec = Convert.ToString(y["CliPec"]).Trim();
                string sCliSdi = Convert.ToString(y["CliSdi"]).Trim();

                string sMsg = "";
                if (sCliPiv == "" && sCliCfi == "")
                    sMsg += "P.IVA mancante o Codice fiscale";
                if (sCliRag == "")
                    sMsg += "R.Sociale mancante ";
                if (sCliInd == "")
                    sMsg += "Indirizzo mancante ";
                if (sCliNcv == "")
                    sMsg += "N.civico mancante ";
                if (sCliCap == "")
                    sMsg += "CAP mancante ";
                if (sCliLoc == "")
                    sMsg += "Località mancante ";
                if (sCliPrv == "")
                    sMsg += "Provincia mancante ";
                if (sCliTip == "E")
                    sMsg += "Cliente estero ";
                if (sCliPec == "" && sCliSdi == "")
                    sMsg += "PEC o Codice destinazione mancanti ";

                // Verifica rapida in memoria
                string docKey = Convert.ToString(y["fat_yfa"]) + "_" + Convert.ToString(y["fat_nfa"]);
                if (dictIvaZeroSenzaNatura.ContainsKey(docKey))
                {
                    sMsg += "IVA " + dictIvaZeroSenzaNatura[docKey] + " ad aliquota ZERO senza codice 'Natura' ";
                }

                // Verifica pagamento in memoria & Fase 3 (Calcolo Scadenza Fine Mese)
                string sTpg = Convert.ToString(y["fat_tpg"]).Trim();
                if (string.IsNullOrEmpty(sTpg))
                {
                    sMsg += "Tipo pagamento non definito ";
                }
                else if (tPagAll != null)
                {
                    DataRow[] rPags = tPagAll.Select("tab_cod='" + sTpg.Replace("'", "''") + "'");
                    if (rPags.Length > 0 && !DBNull.Value.Equals(rPags[0]["tab_fte"]) && ((string)rPags[0]["tab_fte"]).Trim() != "")
                    {
                        string sPCn = Convert.ToString(rPags[0]["tab_ftp"]).Trim();
                        string sPMd = Convert.ToString(rPags[0]["tab_fte"]).Trim();
                        string sPdp = "";

                        string sPar = Convert.ToString(rPags[0]["DscPar"]).Trim();
                        string[] aPar = sPar.Split(',');
                        string sTerm = aPar.Length > 0 ? aPar[0].Trim().ToUpper() : "DF";

                        DateTime dDoc = y["fat_ddo"] != DBNull.Value ? (DateTime)y["fat_ddo"] : DateTime.Today;

                        if (sTerm == "DF")
                        {
                            sPdp = dDoc.ToString("yyyy-MM-dd");
                        }
                        else if (sTerm == "DF30FM")
                        {
                            DateTime dBase = dDoc.AddDays(30);
                            DateTime dFineMese = new DateTime(dBase.Year, dBase.Month, DateTime.DaysInMonth(dBase.Year, dBase.Month));
                            sPdp = dFineMese.ToString("yyyy-MM-dd");
                        }
                        else if (sTerm == "DF60FM")
                        {
                            DateTime dBase = dDoc.AddDays(60);
                            DateTime dFineMese = new DateTime(dBase.Year, dBase.Month, DateTime.DaysInMonth(dBase.Year, dBase.Month));
                            sPdp = dFineMese.ToString("yyyy-MM-dd");
                        }
                        else if (sTerm == "DF90FM")
                        {
                            DateTime dBase = dDoc.AddDays(90);
                            DateTime dFineMese = new DateTime(dBase.Year, dBase.Month, DateTime.DaysInMonth(dBase.Year, dBase.Month));
                            sPdp = dFineMese.ToString("yyyy-MM-dd");
                        }
                        else if (sTerm.StartsWith("DF") && sTerm.EndsWith("FM") && sTerm.Length > 4)
                        {
                            int days = 30;
                            int.TryParse(sTerm.Substring(2, sTerm.Length - 4), out days);
                            DateTime dBase = dDoc.AddDays(days);
                            DateTime dFineMese = new DateTime(dBase.Year, dBase.Month, DateTime.DaysInMonth(dBase.Year, dBase.Month));
                            sPdp = dFineMese.ToString("yyyy-MM-dd");
                        }
                        else if (sTerm.StartsWith("DF") && sTerm.Length > 2)
                        {
                            int days = 0;
                            int.TryParse(sTerm.Substring(2), out days);
                            sPdp = dDoc.AddDays(days).ToString("yyyy-MM-dd");
                        }
                        else
                        {
                            sPdp = dDoc.ToString("yyyy-MM-dd");
                        }

                        if (string.IsNullOrEmpty(sPCn))
                            sMsg += "Condizione pagamento mancante ";
                        if (string.IsNullOrEmpty(sPMd))
                            sMsg += "Modalità pagamento mancante ";
                        if (string.IsNullOrEmpty(sPdp))
                            sMsg += "Data pagamento mancante ";
                        y["FatPag"] = sPCn + "," + sPMd + "," + sPdp + "|";
                    }
                }

                if (sMsg.Length > 150)
                    sMsg = sMsg.Substring(0, 150);
                y["FatMsg"] = sMsg;
            }

            dgv1.DataSource = t;
            lblCnt.Text = t.Rows.Count.ToString("###0");
        }

        private void chkEle_CheckedChanged(object sender, EventArgs e)
        {
            if(chkEle.Checked)
                FillDati("");
            else
                FillDati("DAINV");
        }

        private void cmbYea_SelectionChangeCommitted(object sender, EventArgs e)
        {
            FillDati("");
        }

        private void dgv1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.ColumnIndex == 0)
            {
                CurrencyManager cm = dgv1.BindingContext[dgv1.DataSource, dgv1.DataMember] as CurrencyManager;
                if (cm.Position >= 0)
                {
                    DataRowView r = cm.List[dgv1.CurrentRow.Index] as DataRowView;
                    DataRow x = r.Row;

                    string s = (string)x["FatSel"];
                    if (s == "")
                        x["FatSel"] = "S";
                    else if (s == "S")
                        x["FatSel"] = "";
                }
            }
        }

        private void dgv1_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.ColumnIndex == 2)
            {
                CurrencyManager cm = dgv1.BindingContext[dgv1.DataSource, dgv1.DataMember] as CurrencyManager;
                if (cm.Position >= 0)
                {
                    DataRowView r = cm.List[dgv1.CurrentRow.Index] as DataRowView;
                    DataRow x = r.Row;

                    frmGesDocumento f = new frmGesDocumento();
                    f._strMovFat = "F";
                    f._strCauTpd = "FV";
                    f._strMftYea = (string)x["fat_yfa"];
                    f._strMftNum = (string)x["fat_nfa"];
                    //f._strDocSuf = sSuf;
                    f.ShowDialog();
                }
            }
            else if (e.ColumnIndex == 5)
            {
                CurrencyManager cm = dgv1.BindingContext[dgv1.DataSource, dgv1.DataMember] as CurrencyManager;
                if (cm.Position >= 0)
                {
                    DataRowView r = cm.List[dgv1.CurrentRow.Index] as DataRowView;
                    DataRow x = r.Row;

                    DataTable tTmp = new clsGenTabTmp().TabTmpCli("AnaCli");
                    DataRow k = tTmp.NewRow();
                    k["tmp_cli"] = x["fat_cfo"];
                    k["tmp_cld"] = x["CliRag"];
                    tTmp.Rows.Add(k);

                    frmAnaCliente f = new frmAnaCliente();
                    f._tabTmp = tTmp.Copy();
                    f.ShowDialog();
                }
            }
        }

        private void chkAll_Click(object sender, EventArgs e)
        {
            string sSta = "";
            if (chkAll.Checked)
                sSta = "S";
            DataTable t = (DataTable)dgv1.DataSource;
            foreach (DataRow y in t.Rows)
                y["FatSel"] = sSta;
        }

        private void btnOk_Click(object sender, EventArgs e)
        {
            string sPiv = lblCnfPiv.Text.Trim().ToUpper();                       //Partita IVA
            string sCfi = lblCnfCfi.Text.Trim().ToUpper();                       //Codice Fiscale
            string sCif = lblCnfPiv.Text.Trim().ToUpper();                       //Codice identificativo fiscale (partita IVA?????????????)
            string sMai = lblCnfMai.Text.Trim().ToUpper();                       //Email
            string sRag = lblCnfRag.Text.Trim().ToUpper();                       //Ragione sociale
            string sIba = lblCnfIba.Text.Trim().ToUpper();                       //IBAN
            string sBan = lblCnfBan.Text.Trim().ToUpper();                       //Banca

            string sMsg = "";

            if (sPiv == "")
                sMsg += "Partita IVA emittente mancante" + _clsDef.CRLF;
            //if (sCfi == "")
            //    sMsg += "Codice fiscale mancante" + _clsDef.CRLF;
            if (sCif == "")
                sMsg += "Codice identificativo fiscale mancante (Partita IVA)" + _clsDef.CRLF;
            if (sRag == "")
                sMsg += "Ragione sociale mancante" + _clsDef.CRLF;
            if (sIba == "")
                sMsg += "IBAN mancante" + _clsDef.CRLF;
            if (sBan == "")
                sMsg += "Banca mancante" + _clsDef.CRLF;

            DataView v = new DataView((DataTable)dgv1.DataSource,"","",DataViewRowState.CurrentRows);
            foreach(DataRowView r in v)
            {
                if ((string)r["FatSel"] == "S" && (string)r["FatMsg"] != "")
                    sMsg += (string)r["FatMsg"];
            }

            if(sMsg != "")
                MessageBox.Show(sMsg, "CONTROLLI FATTURA ELETTRONICA", MessageBoxButtons.OK, MessageBoxIcon.Error);
            else
                GenFatElettroniche();
        }
                
        private void GenFatElettroniche()
        {
            DataTable t = (DataTable)dgv1.DataSource;

            DataView v = new DataView(t, "FatSel='S'", "fat_nfa", DataViewRowState.CurrentRows);

            if (v.Count == 0)
            {
                MessageBox.Show("Selezionare i documenti da elaborare!", "CONTROLLI FATTURA ELETTRONICA", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            else
            {
                int iGenCount = 0;
                int iTotCount = v.Count;

                foreach (DataRowView r in v)
                {
                    string sFil = GenFatElettronicheXml(r.Row);
                    if (File.Exists(sFil))
                    {
                        iGenCount++;
                    }
                }

                if (iGenCount > 0)
                {
                    MessageBox.Show("Elaborazione completata con successo!\n\nDocumenti XML generati: " + iGenCount.ToString() + " di " + iTotCount.ToString() + "\nCartella: " + _strPath, "FATTURA ELETTRONICA", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    MessageBox.Show("Nessun documento XML generato.", "FATTURA ELETTRONICA", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
        }
        private string GenFatElettronicheXml(DataRow rowFat)
        {
            if (rowFat == null) return "";

            if (!Directory.Exists(_strPath))
                Directory.CreateDirectory(_strPath);
            string sOldDir = Path.Combine(_strPath, "Old");
            if (!Directory.Exists(sOldDir))
                Directory.CreateDirectory(sOldDir);

            string s = "";
            DataRow[] j;
            string[] a = s.Split('X');

            string sNfa = Convert.ToString(rowFat["fat_nfa"]).Trim();
            string sYea = Convert.ToString(rowFat["fat_yfa"]).Trim();
            DateTime dDdo = rowFat["fat_ddo"] != DBNull.Value ? Convert.ToDateTime(rowFat["fat_ddo"]) : DateTime.Today;

            string sSufNdo = "";    //Suffisso numero documento
            string sSufNfe = "";    //Suffisso numeratore fattura elettronica

            if (_strSuf != "")
            {
                a = _strSuf.Split(',');
                if (a.Length > 0)
                {
                    sSufNdo = a[0];
                    if (a.Length > 1)
                        sSufNfe = a[1];
                }
            }

            // Dati azienda
            string sPec = Convert.ToString(rowFat["CliPec"]).Trim();
            string sPiv = lblCnfPiv.Text.Trim().ToUpper();                       //Partita IVA
            string sCfi = lblCnfCfi.Text.Trim().ToUpper();                       //Codice Fiscale
            string sCif = lblCnfPiv.Text.Trim().ToUpper();                       //Codice identificativo fiscale (partita IVA)
            string sMai = lblCnfMai.Text.Trim();                                 //Email
            string sRag = lblCnfRag.Text.Trim();                                 //Ragione sociale
            string sIba = lblCnfIba.Text.Trim().ToUpper();                       //IBAN
            string sBan = lblCnfBan.Text.Trim();                                 //Banca

            if (sCfi == "")
                sCfi = sPiv;

            string[] aInd = lblCnfInd.Text.Split('-');
            string sInd = aInd.Length > 0 ? aInd[0].Trim() : "";
            string sNcv = aInd.Length > 1 ? aInd[1].Trim() : "";
            string sCap = aInd.Length > 2 ? aInd[2].Trim() : "";
            string sLoc = aInd.Length > 3 ? aInd[3].Trim() : "";
            string sPrv = aInd.Length > 4 ? aInd[4].Trim().ToUpper() : "";

            string sNid = "";                               //Numero identificativo Lettera + numeratore per anno (max 5 caratteri per nome file SdI)
            if (Convert.ToString(rowFat["fat_ele"]).Trim() == "")
            {
                int yearNum = 2026;
                int.TryParse(sYea, out yearNum);
                sNid = Convert.ToChar(Math.Max(65, Math.Min(90, yearNum - 1953))).ToString();

                s = _clsFun.NewNum(sYea, clsDefine.enuNumeratori.NumGesFatElettroniche, 4, _strConSql);

                if (sSufNfe != "")
                    sNid += sSufNfe.Trim() + s.Substring(Math.Max(0, s.Length - 3));
                else
                    sNid += s;

                if (sNid.Length > 5)
                    sNid = sNid.Substring(sNid.Length - 5);

                rowFat["fat_ele"] = sNid;
                s = "UPDATE GesFatTestate SET fat_ele='" + sNid + "' WHERE fat_nfa='" + sNfa + "' AND fat_yfa='" + sYea + "'";
                _clsFun.SqlWrite(s, _strConSql);
            }
            else
            {
                sNid = Convert.ToString(rowFat["fat_ele"]).Trim();
                if (sNid.Length > 5)
                    sNid = sNid.Substring(sNid.Length - 5);
            }

            // Dati cliente
            string sCliPiv = Convert.ToString(rowFat["CliPiv"]).Trim().ToUpper();
            string sCliCfi = Convert.ToString(rowFat["CliCfi"]).Trim().ToUpper();
            string sCliRag = Convert.ToString(rowFat["CliRag"]).Trim();
            string sCliInd = Convert.ToString(rowFat["CliInd"]).Trim();
            string sCliNcv = Convert.ToString(rowFat["CliNcv"]).Trim();
            string sCliCap = Convert.ToString(rowFat["CliCap"]).Trim();
            string sCliLoc = Convert.ToString(rowFat["CliLoc"]).Trim();
            string sCliPrv = Convert.ToString(rowFat["CliPrv"]).Trim().ToUpper();

            string sCliCcp = "";
            if (rowFat.Table.Columns.Contains("CliCcp") && !DBNull.Value.Equals(rowFat["CliCcp"]))
                sCliCcp = Convert.ToString(rowFat["CliCcp"]).Trim().ToUpper();      //Codice cedente prestatore

            string sCliSdi = Convert.ToString(rowFat["CliSdi"]).Trim();
            string sCliNaz = Convert.ToString(rowFat["CliNaz"]).Trim().ToUpper();
            if (string.IsNullOrEmpty(sCliNaz))
                sCliNaz = "IT";

            if (sCliNaz != "IT")
            {
                sCliSdi = "XXXXXXX";
                if (string.IsNullOrEmpty(sCliCap)) sCliCap = "99999";
                if (string.IsNullOrEmpty(sCliPrv)) sCliPrv = "EE";
            }
            else if (string.IsNullOrEmpty(sCliSdi))
            {
                sCliSdi = "0000000";
            }

            bool isPA = (sCliSdi.Length == 6 && sCliNaz == "IT");
            string sFormatoTrasmissione = isPA ? "FPA12" : "FPR12";

            string sFatDdo = dDdo.ToString("yyyy-MM-dd");

            string sFatTdo = "TD01";
            string sRowTdo = Convert.ToString(rowFat["fat_tdo"]).Trim();
            string sRowTpd = "";
            if (rowFat.Table.Columns.Contains("fat_tpd"))
                sRowTpd = Convert.ToString(rowFat["fat_tpd"]).Trim();

            if (sRowTdo == "NA" || sRowTpd == "NA")
                sFatTdo = "TD04";
            else if (sRowTpd == "FD" || sRowTdo == "FD")
                sFatTdo = "TD24";

            string sFatNum = Convert.ToString(rowFat["fat_ndo"]).Trim();
            if (_clsFun.Numerico(sFatNum, "0123456789"))
                sFatNum = Convert.ToUInt64(sFatNum).ToString();
            if (sSufNdo != "")
                sFatNum = sFatNum + sSufNdo.Trim();

            string sFil = Path.Combine(_strPath, "IT" + sCif + "_" + sNid + ".xml");

            s = "SELECT * FROM TabIva ORDER BY tab_ali";
            DataTable tIva = _clsFun.FillTabSql("tabIva", s, false, _strConSql);
            tIva.Columns.Add(new DataColumn { DataType = typeof(decimal), ColumnName = "IvaIva", Caption = "Imposta", ReadOnly = false, DefaultValue = 0m });
            tIva.Columns.Add(new DataColumn { DataType = typeof(decimal), ColumnName = "ImpNoi", Caption = "Importo senza IVA", ReadOnly = false, DefaultValue = 0m });
            tIva.Columns.Add(new DataColumn { DataType = typeof(decimal), ColumnName = "ImpIvo", Caption = "Importo Ivato", ReadOnly = false, DefaultValue = 0m });

            /*** Calcolo i valori del documento ***/
            string sIva = "";
            string sIna = "";               //Natura IVA per aliquote a ZER0
            decimal dTotImp = 0;

            s = "SELECT GesMovimenti.*, AnaArticoli.art_sta FROM GesMovimenti ";
            s += "LEFT JOIN AnaArticoli ON GesMovimenti.mov_art = AnaArticoli.art_cod WHERE ";
            s += "mov_ann= 0 AND ";
            s += "mov_nfa = '" + sNfa.Replace("'", "''") + "' AND ";
            s += "mov_yfa='" + sYea.Replace("'", "''") + "' ";
            s += "ORDER BY mov_rfa";
            DataTable tMov = _clsFun.FillTabSql("GesMovimenti", s, false, _strConSql);
            tMov.Columns.Add(new DataColumn { DataType = typeof(string), ColumnName = "MovAli", Caption = "Aliquota IVA", MaxLength = 5, ReadOnly = false, DefaultValue = "" });
            tMov.Columns.Add(new DataColumn { DataType = typeof(string), ColumnName = "MovIna", Caption = "Natura IVA x alq a ZERO", MaxLength = 6, ReadOnly = false, DefaultValue = "" });

            // Calcoli IVA
            ArrayList aryScon = new ArrayList();
            decimal dTotScon = 0;
            Boolean bTotScon = true;

            foreach (DataRow y in tMov.Rows)
            {
                string movIvaCod = Convert.ToString(y["mov_iva"]).Trim();
                j = tIva.Select("tab_cod='" + movIvaCod.Replace("'", "''") + "'");
                if (j.Length > 0)
                {
                    decimal dAli = Convert.ToDecimal(j[0]["tab_ali"]);
                    sIva = dAli.ToString("0.00", CultureInfo.InvariantCulture);
                    sIna = Convert.ToString(j[0]["tab_ele"]).Trim();
                    decimal dImpVal = !DBNull.Value.Equals(y["mov_imp"]) ? Convert.ToDecimal(y["mov_imp"]) : 0m;
                    j[0]["ImpNoi"] = Convert.ToDecimal(j[0]["ImpNoi"]) + dImpVal;
                    y["MovAli"] = sIva;
                    y["MovIna"] = sIna;

                    if (!DBNull.Value.Equals(y["mov_ori"]) && Convert.ToString(y["mov_ori"]).Trim() != "")
                    {
                        string sOri = Convert.ToString(y["mov_ori"]);
                        string[] aOri = sOri.Split(',');
                        if (aOri.Length > 4)
                        {
                            string sKey = sOri.Length >= 24 ? sOri.Substring(0, 24) : sOri;
                            if (aryScon.IndexOf(sKey) < 0)
                            {
                                aryScon.Add(sKey);
                                if (aOri.Length > 5)
                                {
                                    decimal dSconVal = 0m;
                                    decimal.TryParse(aOri[5].Replace(".", ","), out dSconVal);
                                    dTotScon += dSconVal;
                                }
                            }
                        }
                    }
                    else if (dImpVal > 0)
                    {
                        bTotScon = false;
                    }
                }
            }

            if (!bTotScon)
                dTotScon = 0;

            dTotImp = 0;

            DataTable tIvaOk = tIva.Clone();
            foreach (DataRow y in tIva.Rows)
            {
                decimal dImpNoi = Convert.ToDecimal(y["ImpNoi"]);
                if (dImpNoi != 0)
                {
                    decimal dAli = Convert.ToDecimal(y["tab_ali"]);
                    decimal dIvaVal = Math.Round(_clsFun.ValIva(dImpNoi, dAli), 2, MidpointRounding.ToEven);
                    decimal dImpIvo = Math.Round(dImpNoi + dIvaVal, 2, MidpointRounding.ToEven);
                    dImpNoi = dImpIvo - dIvaVal;

                    y["IvaIva"] = dIvaVal;
                    y["ImpIvo"] = dImpIvo;
                    y["ImpNoi"] = dImpNoi;
                    dTotImp += dImpIvo;

                    tIvaOk.ImportRow(y);
                }
            }

            tIva = tIvaOk.Copy();

            if (dTotScon > 0 && dTotImp != dTotScon)
            {
                dTotImp = 0;
                foreach (DataRow y in tIva.Rows)
                {
                    dTotImp += Convert.ToDecimal(y["ImpNoi"]) + Convert.ToDecimal(y["IvaIva"]);
                }
            }

            // Scrittura File XML con codifica UTF-8 garantita
            using (StreamWriter sw = new StreamWriter(sFil, false, new UTF8Encoding(false)))
            {
                sw.Write("<?xml version=\"1.0\" encoding=\"utf-8\"?>" + _clsDef.CRLF);
                sw.Write("<p:FatturaElettronica xmlns:ds=\"http://www.w3.org/2000/09/xmldsig#\" xmlns:xis=\"http://www.w3.org/2001/XMLSchema-instance\" versione=\"" + sFormatoTrasmissione + "\" xmlns:p=\"http://ivaservizi.agenziaentrate.gov.it/docs/xsd/fatture/v1.2\">" + _clsDef.CRLF);

                /*** TESTATA INIZIO ***/
                sw.Write("    <FatturaElettronicaHeader>" + _clsDef.CRLF);

                sw.Write("        <DatiTrasmissione>" + _clsDef.CRLF);
                sw.Write("            <IdTrasmittente>" + _clsDef.CRLF);
                sw.Write("                <IdPaese>IT</IdPaese>" + _clsDef.CRLF);
                sw.Write("                <IdCodice>" + XmlEsc(sCif) + "</IdCodice>" + _clsDef.CRLF);
                sw.Write("            </IdTrasmittente>" + _clsDef.CRLF);
                sw.Write("            <ProgressivoInvio>" + XmlEsc(sNid) + "</ProgressivoInvio>" + _clsDef.CRLF);
                sw.Write("            <FormatoTrasmissione>" + sFormatoTrasmissione + "</FormatoTrasmissione>" + _clsDef.CRLF);
                sw.Write("            <CodiceDestinatario>" + XmlEsc(sCliSdi) + "</CodiceDestinatario>" + _clsDef.CRLF);
                if (!string.IsNullOrEmpty(sPec) && sCliSdi == "0000000")
                    sw.Write("            <PECDestinatario>" + XmlEsc(sPec) + "</PECDestinatario>" + _clsDef.CRLF);
                sw.Write("        </DatiTrasmissione>" + _clsDef.CRLF);

                /*** CedentePrestatore INIZIO ***/
                sw.Write("        <CedentePrestatore>" + _clsDef.CRLF);
                sw.Write("            <DatiAnagrafici>" + _clsDef.CRLF);
                sw.Write("                <IdFiscaleIVA>" + _clsDef.CRLF);
                sw.Write("                    <IdPaese>IT</IdPaese>" + _clsDef.CRLF);
                sw.Write("                    <IdCodice>" + XmlEsc(sPiv) + "</IdCodice>" + _clsDef.CRLF);
                sw.Write("                </IdFiscaleIVA>" + _clsDef.CRLF);
                if (!string.IsNullOrEmpty(sCfi))
                    sw.Write("                <CodiceFiscale>" + XmlEsc(sCfi) + "</CodiceFiscale>" + _clsDef.CRLF);
                sw.Write("                <Anagrafica>" + _clsDef.CRLF);
                sw.Write("                    <Denominazione>" + XmlEsc(sRag) + "</Denominazione>" + _clsDef.CRLF);
                sw.Write("                </Anagrafica>" + _clsDef.CRLF);
                sw.Write("                <RegimeFiscale>RF01</RegimeFiscale>" + _clsDef.CRLF);
                sw.Write("            </DatiAnagrafici>" + _clsDef.CRLF);
                sw.Write("            <Sede>" + _clsDef.CRLF);
                sw.Write("                <Indirizzo>" + XmlEsc(sInd) + "</Indirizzo>" + _clsDef.CRLF);
                if (!string.IsNullOrEmpty(sNcv))
                    sw.Write("                <NumeroCivico>" + XmlEsc(sNcv) + "</NumeroCivico>" + _clsDef.CRLF);
                sw.Write("                <CAP>" + XmlEsc(sCap) + "</CAP>" + _clsDef.CRLF);
                sw.Write("                <Comune>" + XmlEsc(sLoc) + "</Comune>" + _clsDef.CRLF);
                if (!string.IsNullOrEmpty(sPrv))
                    sw.Write("                <Provincia>" + XmlEsc(sPrv) + "</Provincia>" + _clsDef.CRLF);
                sw.Write("                <Nazione>IT</Nazione>" + _clsDef.CRLF);
                sw.Write("            </Sede>" + _clsDef.CRLF);
                if (!string.IsNullOrEmpty(sCliCcp))
                    sw.Write("            <RiferimentoAmministrazione>" + XmlEsc(sCliCcp) + "</RiferimentoAmministrazione>" + _clsDef.CRLF);

                sw.Write("        </CedentePrestatore>" + _clsDef.CRLF);
                /*** CedentePrestatore FINE ***/

                string sTot = dTotImp.ToString("0.00", CultureInfo.InvariantCulture);

                sw.Write("        <CessionarioCommittente>" + _clsDef.CRLF);
                sw.Write("            <DatiAnagrafici>" + _clsDef.CRLF);
                if (!string.IsNullOrEmpty(sCliPiv))
                {
                    sw.Write("                <IdFiscaleIVA>" + _clsDef.CRLF);
                    sw.Write("                    <IdPaese>" + XmlEsc(sCliNaz) + "</IdPaese>" + _clsDef.CRLF);
                    sw.Write("                    <IdCodice>" + XmlEsc(sCliPiv) + "</IdCodice>" + _clsDef.CRLF);
                    sw.Write("                </IdFiscaleIVA>" + _clsDef.CRLF);
                }
                if (!string.IsNullOrEmpty(sCliCfi))
                    sw.Write("                <CodiceFiscale>" + XmlEsc(sCliCfi) + "</CodiceFiscale>" + _clsDef.CRLF);
                sw.Write("                <Anagrafica>" + _clsDef.CRLF);
                sw.Write("                    <Denominazione>" + XmlEsc(sCliRag) + "</Denominazione>" + _clsDef.CRLF);
                sw.Write("                </Anagrafica>" + _clsDef.CRLF);
                sw.Write("            </DatiAnagrafici>" + _clsDef.CRLF);
                sw.Write("            <Sede>" + _clsDef.CRLF);
                sw.Write("                <Indirizzo>" + XmlEsc(sCliInd) + "</Indirizzo>" + _clsDef.CRLF);
                if (!string.IsNullOrEmpty(sCliNcv))
                    sw.Write("                <NumeroCivico>" + XmlEsc(sCliNcv) + "</NumeroCivico>" + _clsDef.CRLF);
                sw.Write("                <CAP>" + XmlEsc(sCliCap) + "</CAP>" + _clsDef.CRLF);
                sw.Write("                <Comune>" + XmlEsc(sCliLoc) + "</Comune>" + _clsDef.CRLF);
                if (!string.IsNullOrEmpty(sCliPrv) && sCliNaz == "IT")
                    sw.Write("                <Provincia>" + XmlEsc(sCliPrv) + "</Provincia>" + _clsDef.CRLF);
                sw.Write("                <Nazione>" + XmlEsc(sCliNaz) + "</Nazione>" + _clsDef.CRLF);
                sw.Write("            </Sede>" + _clsDef.CRLF);
                sw.Write("        </CessionarioCommittente>" + _clsDef.CRLF);
                sw.Write("    </FatturaElettronicaHeader>" + _clsDef.CRLF);

                sw.Write("    <FatturaElettronicaBody>" + _clsDef.CRLF);

                // Ricerca eventuali DDT collegati per fatture differite (TD24)
                DataTable tDdt = null;
                if (sFatTdo == "TD24" || sRowTpd == "FD")
                {
                    s = "SELECT mot_ndo, mot_ddo FROM GesMovTestate WHERE mot_yfa='" + sYea.Replace("'", "''") + "' AND mot_nfa='" + sNfa.Replace("'", "''") + "' AND mot_ndo <> '' AND mot_ann = 0 GROUP BY mot_ndo, mot_ddo ORDER BY mot_ddo, mot_ndo";
                    tDdt = _clsFun.FillTabSql("GesMovTestateDDT", s, false, _strConSql);
                }

                sw.Write("        <DatiGenerali>" + _clsDef.CRLF);
                sw.Write("            <DatiGeneraliDocumento>" + _clsDef.CRLF);
                sw.Write("                <TipoDocumento>" + XmlEsc(sFatTdo) + "</TipoDocumento>" + _clsDef.CRLF);
                sw.Write("                <Divisa>EUR</Divisa>" + _clsDef.CRLF);
                sw.Write("                <Data>" + XmlEsc(sFatDdo) + "</Data>" + _clsDef.CRLF);
                sw.Write("                <Numero>" + XmlEsc(sFatNum) + "</Numero>" + _clsDef.CRLF);
                sw.Write("                <ImportoTotaleDocumento>" + sTot + "</ImportoTotaleDocumento>" + _clsDef.CRLF);
                sw.Write("                <Causale>VENDITA</Causale>" + _clsDef.CRLF);
                sw.Write("            </DatiGeneraliDocumento>" + _clsDef.CRLF);

                if (tDdt != null && tDdt.Rows.Count > 0)
                {
                    foreach (DataRow rDdt in tDdt.Rows)
                    {
                        string sNumDdt = Convert.ToString(rDdt["mot_ndo"]).Trim();
                        string sDatDdt = "";
                        if (rDdt["mot_ddo"] != DBNull.Value)
                            sDatDdt = ((DateTime)rDdt["mot_ddo"]).ToString("yyyy-MM-dd");

                        if (!string.IsNullOrEmpty(sNumDdt) && !string.IsNullOrEmpty(sDatDdt))
                        {
                            sw.Write("            <DatiDDT>" + _clsDef.CRLF);
                            sw.Write("                <NumeroDDT>" + XmlEsc(sNumDdt) + "</NumeroDDT>" + _clsDef.CRLF);
                            sw.Write("                <DataDDT>" + XmlEsc(sDatDdt) + "</DataDDT>" + _clsDef.CRLF);
                            sw.Write("            </DatiDDT>" + _clsDef.CRLF);
                        }
                    }
                }

                sw.Write("        </DatiGenerali>" + _clsDef.CRLF);
                sw.Write("        <DatiBeniServizi>" + _clsDef.CRLF);

                int i = 0;
                foreach (DataRow y in tMov.Rows)
                {
                    decimal dMovImp = !DBNull.Value.Equals(y["mov_imp"]) ? Convert.ToDecimal(y["mov_imp"]) : 0m;
                    string sArd = Convert.ToString(y["mov_ard"]).Trim();

                    if (dMovImp == 0)
                    {
                        if (sArd != "")
                        {
                            sIva = "0.00";
                            sIna = "N2.2";
                            j = tIva.Select("tab_ali > 0");
                            if (j.Length > 0)
                            {
                                sIva = Convert.ToDecimal(j[0]["tab_ali"]).ToString("0.00", CultureInfo.InvariantCulture);
                                sIna = "";
                            }
                            else if (tIva.Rows.Count > 0)
                            {
                                sIva = Convert.ToDecimal(tIva.Rows[0]["tab_ali"]).ToString("0.00", CultureInfo.InvariantCulture);
                                sIna = Convert.ToString(tIva.Rows[0]["tab_ele"]).Trim();
                                if (string.IsNullOrEmpty(sIna)) sIna = "N2.2";
                            }

                            i++;
                            string sRfa = i.ToString();

                            sw.Write("            <DettaglioLinee>" + _clsDef.CRLF);
                            sw.Write("                <NumeroLinea>" + sRfa + "</NumeroLinea>" + _clsDef.CRLF);
                            sw.Write("                <Descrizione>" + XmlEsc(sArd) + "</Descrizione>" + _clsDef.CRLF);
                            sw.Write("                <PrezzoUnitario>0.00</PrezzoUnitario>" + _clsDef.CRLF);
                            sw.Write("                <PrezzoTotale>0.00</PrezzoTotale>" + _clsDef.CRLF);
                            sw.Write("                <AliquotaIVA>" + sIva + "</AliquotaIVA>" + _clsDef.CRLF);
                            if (Convert.ToDecimal(sIva, CultureInfo.InvariantCulture) == 0 && !string.IsNullOrEmpty(sIna))
                                sw.Write("                <Natura>" + XmlEsc(sIna) + "</Natura>" + _clsDef.CRLF);

                            sw.Write("            </DettaglioLinee>" + _clsDef.CRLF);
                        }
                    }
                    else
                    {
                        i++;
                        string sRfa = i.ToString();
                        string sArt = Convert.ToString(y["mov_art"]).Trim();
                        string sSta = Convert.ToString(y["art_sta"]).Trim();

                        string sUmi = Convert.ToString(y["mov_umi"]).Trim();
                        if (string.IsNullOrEmpty(sUmi))
                            sUmi = "NR";

                        decimal dQta = !DBNull.Value.Equals(y["mov_qta"]) ? Convert.ToDecimal(y["mov_qta"]) : 1m;
                        if (sUmi == "KG" && sSta != "R" && !DBNull.Value.Equals(y["mov_qkg"]) && Convert.ToDecimal(y["mov_qkg"]) > 0 && Convert.ToDecimal(y["mov_qkg"]) < 100)
                            dQta = Convert.ToDecimal(y["mov_qkg"]);

                        if (dQta == 0m) dQta = 1m;

                        decimal dPrz = dMovImp / dQta;
                        string sQta = dQta.ToString("0.000", CultureInfo.InvariantCulture);
                        string sPrz = dPrz.ToString("0.000000", CultureInfo.InvariantCulture);
                        string sImp = dMovImp.ToString("0.00", CultureInfo.InvariantCulture);

                        sIva = Convert.ToString(y["MovAli"]).Replace(",", ".");
                        sIna = Convert.ToString(y["MovIna"]).Trim();
                        decimal dAliVal = 0m;
                        decimal.TryParse(sIva, NumberStyles.Any, CultureInfo.InvariantCulture, out dAliVal);

                        sw.Write("            <DettaglioLinee>" + _clsDef.CRLF);
                        sw.Write("                <NumeroLinea>" + sRfa + "</NumeroLinea>" + _clsDef.CRLF);
                        sw.Write("                <Descrizione>" + XmlEsc(sArd) + "</Descrizione>" + _clsDef.CRLF);
                        sw.Write("                <Quantita>" + sQta + "</Quantita>" + _clsDef.CRLF);
                        sw.Write("                <UnitaMisura>" + XmlEsc(sUmi) + "</UnitaMisura>" + _clsDef.CRLF);
                        sw.Write("                <PrezzoUnitario>" + sPrz + "</PrezzoUnitario>" + _clsDef.CRLF);
                        sw.Write("                <PrezzoTotale>" + sImp + "</PrezzoTotale>" + _clsDef.CRLF);
                        sw.Write("                <AliquotaIVA>" + dAliVal.ToString("0.00", CultureInfo.InvariantCulture) + "</AliquotaIVA>" + _clsDef.CRLF);
                        if (dAliVal == 0)
                        {
                            if (string.IsNullOrEmpty(sIna)) sIna = "N2.2";
                            sw.Write("                <Natura>" + XmlEsc(sIna) + "</Natura>" + _clsDef.CRLF);
                        }

                        sw.Write("            </DettaglioLinee>" + _clsDef.CRLF);
                    }
                }

                sIva = "";
                decimal dIva = 0;

                foreach (DataRow y in tIva.Rows)
                {
                    decimal dImpIvo = Convert.ToDecimal(y["ImpIvo"]);
                    decimal dImpNoi = Convert.ToDecimal(y["ImpNoi"]);
                    if (dImpIvo > 0 || dImpNoi > 0)
                    {
                        decimal dAli = Convert.ToDecimal(y["tab_ali"]);
                        sIva = dAli.ToString("0.00", CultureInfo.InvariantCulture);
                        sIna = Convert.ToString(y["tab_ele"]).Trim();
                        string sImp = dImpNoi.ToString("0.00", CultureInfo.InvariantCulture);

                        dIva = _clsFun.ValIva(dImpNoi, dAli);
                        string sVal = dIva.ToString("0.00", CultureInfo.InvariantCulture);

                        sw.Write("            <DatiRiepilogo>" + _clsDef.CRLF);
                        sw.Write("                <AliquotaIVA>" + sIva + "</AliquotaIVA>" + _clsDef.CRLF);
                        if (dAli == 0)
                        {
                            if (string.IsNullOrEmpty(sIna)) sIna = "N2.2";
                            sw.Write("                <Natura>" + XmlEsc(sIna) + "</Natura>" + _clsDef.CRLF);
                            sw.Write("                <RiferimentoNormativo>Operazione non imponibile / esente</RiferimentoNormativo>" + _clsDef.CRLF);
                        }
                        sw.Write("                <ImponibileImporto>" + sImp + "</ImponibileImporto>" + _clsDef.CRLF);
                        sw.Write("                <Imposta>" + sVal + "</Imposta>" + _clsDef.CRLF);
                        if (dAli > 0)
                        {
                            if (isPA)
                                sw.Write("                <EsigibilitaIVA>S</EsigibilitaIVA>" + _clsDef.CRLF);
                            else
                                sw.Write("                <EsigibilitaIVA>I</EsigibilitaIVA>" + _clsDef.CRLF);
                        }
                        sw.Write("            </DatiRiepilogo>" + _clsDef.CRLF);
                    }
                }

                sw.Write("        </DatiBeniServizi>" + _clsDef.CRLF);

                string sFatPag = Convert.ToString(rowFat["FatPag"]).Trim();
                if (!string.IsNullOrEmpty(sFatPag))
                {
                    string[] aPagBlocks = sFatPag.Split(new char[] { '|' }, StringSplitOptions.RemoveEmptyEntries);
                    foreach (string ss in aPagBlocks)
                    {
                        if (ss.Trim() != "")
                        {
                            string[] aa = ss.Split(',');
                            string sFatPCn = aa.Length > 0 ? aa[0].Trim() : "TP02";
                            string sFatPMd = aa.Length > 1 ? aa[1].Trim() : "MP01";
                            string sFatPdp = aa.Length > 2 ? aa[2].Trim() : sFatDdo;

                            sw.Write("        <DatiPagamento>" + _clsDef.CRLF);
                            sw.Write("            <CondizioniPagamento>" + XmlEsc(sFatPCn) + "</CondizioniPagamento>" + _clsDef.CRLF);
                            sw.Write("            <DettaglioPagamento>" + _clsDef.CRLF);
                            sw.Write("                <ModalitaPagamento>" + XmlEsc(sFatPMd) + "</ModalitaPagamento>" + _clsDef.CRLF);
                            sw.Write("                <DataScadenzaPagamento>" + XmlEsc(sFatPdp) + "</DataScadenzaPagamento>" + _clsDef.CRLF);
                            sw.Write("                <ImportoPagamento>" + sTot + "</ImportoPagamento>" + _clsDef.CRLF);
                            if (sFatPMd == "MP05")
                            {
                                if (!string.IsNullOrEmpty(sBan))
                                    sw.Write("                <IstitutoFinanziario>" + XmlEsc(sBan) + "</IstitutoFinanziario>" + _clsDef.CRLF);
                                if (!string.IsNullOrEmpty(sIba))
                                    sw.Write("                <IBAN>" + XmlEsc(sIba) + "</IBAN>" + _clsDef.CRLF);
                            }
                            sw.Write("            </DettaglioPagamento>" + _clsDef.CRLF);
                            sw.Write("        </DatiPagamento>" + _clsDef.CRLF);
                        }
                    }
                }
                else
                {
                    // Blocco di pagamento predefinito se non specificato
                    sw.Write("        <DatiPagamento>" + _clsDef.CRLF);
                    sw.Write("            <CondizioniPagamento>TP02</CondizioniPagamento>" + _clsDef.CRLF);
                    sw.Write("            <DettaglioPagamento>" + _clsDef.CRLF);
                    sw.Write("                <ModalitaPagamento>MP01</ModalitaPagamento>" + _clsDef.CRLF);
                    sw.Write("                <DataScadenzaPagamento>" + XmlEsc(sFatDdo) + "</DataScadenzaPagamento>" + _clsDef.CRLF);
                    sw.Write("                <ImportoPagamento>" + sTot + "</ImportoPagamento>" + _clsDef.CRLF);
                    sw.Write("            </DettaglioPagamento>" + _clsDef.CRLF);
                    sw.Write("        </DatiPagamento>" + _clsDef.CRLF);
                }

                sw.Write("    </FatturaElettronicaBody>" + _clsDef.CRLF);
                sw.Write("</p:FatturaElettronica>" + _clsDef.CRLF);
            }

            return sFil;
        }

        private string FatEleInvio(string strFil)
        {
            string s = "";
            string sFil = _strPath + "AgenziaEntrate.ini";
            string sRig = "";
            string sMsg = "";

            if (File.Exists(sFil))
            {
                string sSmpt = "";
                string sMitt = lblCnfMai.Text;
                string sMittName = lblCnfRag.Text;
                string sPswd = "";
                string sDest = lblCnfAgp.Text;
                string sDestName = "AGENZIA DELLE ENTRATE";
                string sDsCC = "";                              //Copia carbone
                int iPort = 25;

                if (!File.Exists(sFil))
                    MessageBox.Show("File ini di configurazione non presente!", "INVIO FATTURA ELETTRONICA", MessageBoxButtons.OK, MessageBoxIcon.Error);
                else
                {
                    using (StreamReader sr = new StreamReader(sFil, System.Text.Encoding.Default))
                    {
                        while ((sRig = sr.ReadLine()) != null)
                        {
                            if (sRig.Length > 6)
                            {
                                if (sRig.Length > 10 && sRig.Substring(0, 4) == "SMTP")
                                    sSmpt = sRig.Substring(5).Trim();
                                else if (sRig.Length > 10 && sRig.Substring(0, 4) == "PSWD")
                                    sPswd = sRig.Substring(5).Trim();
                                else if (sRig.Length > 10 && sRig.Substring(0, 4) == "DSCC")
                                    sDsCC = sRig.Substring(5).Trim();
                                else if (sRig.Length > 6 && sRig.Substring(0, 4) == "PORT")
                                    iPort = Convert.ToInt16(sRig.Substring(5).Trim());
                            }
                        }
                    }
                }

                if (sSmpt == "")
                    sMsg += "SMTP non definto! " + _clsDef.CRLF;
                if (sPswd == "")
                    sMsg += "Password non definta! " + _clsDef.CRLF;
                if (sMitt == "")
                    sMsg += "Mittente non definto! " + _clsDef.CRLF;
                if (sDest == "")
                    sMsg += "Destinatario non definto! " + _clsDef.CRLF;

                if (sMsg != "")
                    MessageBox.Show(sMsg, "PROBLEMI INVIO FATTURA ELETTRONICA", MessageBoxButtons.OK, MessageBoxIcon.Error);
                else
                {
                    string sOld2 = _strPath + "Old\\" + Path.GetFileName(strFil) + "_" + DateTime.Now.ToString("yyyyMMddHHmmss");
                    File.Copy(strFil, sOld2);

                    clsMail _clsMail = new clsMail();
                    _clsMail._strSmptCli = sSmpt;
                    _clsMail._strMittInd = sMitt;
                    _clsMail._strMittNome = sMittName;
                    _clsMail._strMittPwd = sPswd;
                    _clsMail._strDestInd = sDest;
                    _clsMail._strDestNome = sDestName;
                    _clsMail._strCopiaCarbone = sDsCC;
                    _clsMail._intPorta = iPort;

                    _clsMail._strObj = "Fattura " + Path.GetFileName(strFil);
                    _clsMail._strBody = "Fattura " + Path.GetFileName(strFil);
                    _clsMail._strAttachment = strFil;

                    s = _clsMail.InvioMailAG();


                    if (s != "")
                        MessageBox.Show(s, "PROBLEMI INVIO FATTURA ELETTRONICA", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    else
                    {
                        string sOld = _strPath + "Old\\" + Path.GetFileName(strFil) + "_" + DateTime.Now.ToString("yyyyMMddHHmmss");

                        Console.WriteLine("xxxx");

                        File.Move(strFil, sOld);
                    }
                }

                //string sMai = (string)t.Rows[0]["tab_mai"];

                //a = sMai.Split(',');

                //if (a.Length < 4)
                //    MessageBox.Show("Parametri per invio mail mancanti in tabella fornitori!", "INVIO ORDINE VIA MAIL", MessageBoxButtons.OK, MessageBoxIcon.Error);
                //else
                //{

                //    string sSmptCli = a[0].Trim();
                //    string sMitUsr = a[1].Trim();
                //    string sMitPwd = a[2].Trim();
                //    string sMaiDest = a[3].Trim();

                //    string sMaiCC = "";
                //    if (a.Length > 4)
                //        sMaiCC = a[4].Trim();


                //    //_clsMail._strCopiaCarbone += "psecchettin@tiscali.it";

                //    //_clsMail._strObj = "Ordine O110109.010 - 13/03/2012 - 14:16";


                //    if (s != "")
                //        MessageBox.Show(s, "ERRORE SU INVIO ORDINE VIA MAIL", MessageBoxButtons.OK, MessageBoxIcon.Error);
                //    else
                //    {
                //        cmbOftSta.SelectedValue = _clsDef.DIVDIV;
                //        MessageBox.Show("ORDINE INVIATO", "CONTROLLO INVIO ORDINE VIA MAIL", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                //    }
                //}

            }

            return sMsg;
        }

    }
}
