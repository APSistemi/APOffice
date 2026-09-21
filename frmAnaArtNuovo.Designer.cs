namespace APOffice
{
    partial class frmAnaArtNuovo
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.menuStrip1 = new System.Windows.Forms.MenuStrip();
            this.esciToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.panelHeader = new System.Windows.Forms.Panel();
            this.lblHeaderSub = new System.Windows.Forms.Label();
            this.lblHeaderTitle = new System.Windows.Forms.Label();
            this.panelFooter = new System.Windows.Forms.Panel();
            this.btnSalvaNuovo = new System.Windows.Forms.Button();
            this.btnOk = new System.Windows.Forms.Button();
            this.btnEsci = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.txtArtCod = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.txtArtDes = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.label5 = new System.Windows.Forms.Label();
            this.label6 = new System.Windows.Forms.Label();
            this.label7 = new System.Windows.Forms.Label();
            this.label8 = new System.Windows.Forms.Label();
            this.label9 = new System.Windows.Forms.Label();
            this.label10 = new System.Windows.Forms.Label();
            this.label11 = new System.Windows.Forms.Label();
            this.cmbArtIva = new System.Windows.Forms.ComboBox();
            this.cmbArtUmi = new System.Windows.Forms.ComboBox();
            this.cmbArtTgr = new System.Windows.Forms.ComboBox();
            this.txtArtNet = new System.Windows.Forms.TextBox();
            this.cmbArtEc3 = new System.Windows.Forms.ComboBox();
            this.cmbArtRep = new System.Windows.Forms.ComboBox();
            this.txtEanEan = new System.Windows.Forms.TextBox();
            this.txtLiaPrc = new System.Windows.Forms.TextBox();
            this.txtLivPrv = new System.Windows.Forms.TextBox();
            this.cmbLiaFor = new System.Windows.Forms.ComboBox();
            this.cmbLivLis = new System.Windows.Forms.ComboBox();
            this.chkMem = new System.Windows.Forms.CheckBox();
            this.chkRicercaAuto = new System.Windows.Forms.CheckBox();
            this.btnCercaWeb = new System.Windows.Forms.Button();
            this.lblWebStatus = new System.Windows.Forms.Label();
            this.picWebProduct = new System.Windows.Forms.PictureBox();
            this.lblPicCaption = new System.Windows.Forms.Label();
            this.lblArtArf = new System.Windows.Forms.Label();
            this.txtArtArf = new System.Windows.Forms.TextBox();
            this.lblCosto = new System.Windows.Forms.Label();
            this.lblRicarico = new System.Windows.Forms.Label();
            this.txtRicarico = new System.Windows.Forms.TextBox();
            this.lblPrezzo = new System.Windows.Forms.Label();
            this.panelHeader.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.picWebProduct)).BeginInit();
            this.menuStrip1.SuspendLayout();
            this.SuspendLayout();
            // 
            // menuStrip1
            // 
            this.menuStrip1.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(241)))), ((int)(((byte)(245)))), ((int)(((byte)(249)))));
            this.menuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.esciToolStripMenuItem});
            this.menuStrip1.Location = new System.Drawing.Point(0, 0);
            this.menuStrip1.Name = "menuStrip1";
            this.menuStrip1.Size = new System.Drawing.Size(840, 24);
            this.menuStrip1.TabIndex = 0;
            this.menuStrip1.Text = "menuStrip1";
            // 
            // esciToolStripMenuItem
            // 
            this.esciToolStripMenuItem.Name = "esciToolStripMenuItem";
            this.esciToolStripMenuItem.Size = new System.Drawing.Size(39, 20);
            this.esciToolStripMenuItem.Text = "Esci";
            this.esciToolStripMenuItem.Click += new System.EventHandler(this.esciToolStripMenuItem_Click);
            // 
            // panelHeader
            // 
            this.panelHeader.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(30)))), ((int)(((byte)(41)))), ((int)(((byte)(59)))));
            this.panelHeader.Controls.Add(this.lblHeaderSub);
            this.panelHeader.Controls.Add(this.lblHeaderTitle);
            this.panelHeader.Dock = System.Windows.Forms.DockStyle.Top;
            this.panelHeader.Location = new System.Drawing.Point(0, 24);
            this.panelHeader.Name = "panelHeader";
            this.panelHeader.Size = new System.Drawing.Size(840, 56);
            this.panelHeader.TabIndex = 35;
            // 
            // lblHeaderSub
            // 
            this.lblHeaderSub.AutoSize = true;
            this.lblHeaderSub.Font = new System.Drawing.Font("Segoe UI", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblHeaderSub.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(148)))), ((int)(((byte)(163)))), ((int)(((byte)(184)))));
            this.lblHeaderSub.Location = new System.Drawing.Point(16, 32);
            this.lblHeaderSub.Name = "lblHeaderSub";
            this.lblHeaderSub.Size = new System.Drawing.Size(434, 15);
            this.lblHeaderSub.TabIndex = 1;
            this.lblHeaderSub.Text = "Ricerca automatica Web EAN & compilazione veloce informazioni prodotto da Internet";
            // 
            // lblHeaderTitle
            // 
            this.lblHeaderTitle.AutoSize = true;
            this.lblHeaderTitle.Font = new System.Drawing.Font("Segoe UI", 13.8F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblHeaderTitle.ForeColor = System.Drawing.Color.White;
            this.lblHeaderTitle.Location = new System.Drawing.Point(14, 6);
            this.lblHeaderTitle.Name = "lblHeaderTitle";
            this.lblHeaderTitle.Size = new System.Drawing.Size(325, 25);
            this.lblHeaderTitle.TabIndex = 0;
            this.lblHeaderTitle.Text = "🛍️ INSERIMENTO ARTICOLO VELOCE";
            // 
            // panelFooter
            // 
            this.panelFooter.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(15)))), ((int)(((byte)(23)))), ((int)(((byte)(42)))));
            this.panelFooter.Controls.Add(this.btnSalvaNuovo);
            this.panelFooter.Controls.Add(this.btnOk);
            this.panelFooter.Controls.Add(this.btnEsci);
            this.panelFooter.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.panelFooter.Location = new System.Drawing.Point(0, 595);
            this.panelFooter.Name = "panelFooter";
            this.panelFooter.Size = new System.Drawing.Size(840, 65);
            this.panelFooter.TabIndex = 36;
            // 
            // btnSalvaNuovo
            // 
            this.btnSalvaNuovo.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(5)))), ((int)(((byte)(150)))), ((int)(((byte)(105)))));
            this.btnSalvaNuovo.Cursor = System.Windows.Forms.Cursors.Hand;
            this.btnSalvaNuovo.FlatAppearance.BorderSize = 0;
            this.btnSalvaNuovo.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnSalvaNuovo.Font = new System.Drawing.Font("Segoe UI", 11.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnSalvaNuovo.ForeColor = System.Drawing.Color.White;
            this.btnSalvaNuovo.Location = new System.Drawing.Point(14, 9);
            this.btnSalvaNuovo.Name = "btnSalvaNuovo";
            this.btnSalvaNuovo.Size = new System.Drawing.Size(290, 46);
            this.btnSalvaNuovo.TabIndex = 27;
            this.btnSalvaNuovo.Text = "✔ Salva e Continua (F10)";
            this.btnSalvaNuovo.UseVisualStyleBackColor = false;
            this.btnSalvaNuovo.Click += new System.EventHandler(this.btnSalvaNuovo_Click);
            // 
            // btnOk
            // 
            this.btnOk.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(37)))), ((int)(((byte)(99)))), ((int)(((byte)(235)))));
            this.btnOk.Cursor = System.Windows.Forms.Cursors.Hand;
            this.btnOk.FlatAppearance.BorderSize = 0;
            this.btnOk.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnOk.Font = new System.Drawing.Font("Segoe UI", 11.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnOk.ForeColor = System.Drawing.Color.White;
            this.btnOk.Location = new System.Drawing.Point(314, 9);
            this.btnOk.Name = "btnOk";
            this.btnOk.Size = new System.Drawing.Size(250, 46);
            this.btnOk.TabIndex = 28;
            this.btnOk.Text = "💾 Salva e Chiudi";
            this.btnOk.UseVisualStyleBackColor = false;
            this.btnOk.Click += new System.EventHandler(this.btnOk_Click);
            // 
            // btnEsci
            // 
            this.btnEsci.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(220)))), ((int)(((byte)(38)))), ((int)(((byte)(38)))));
            this.btnEsci.Cursor = System.Windows.Forms.Cursors.Hand;
            this.btnEsci.FlatAppearance.BorderSize = 0;
            this.btnEsci.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnEsci.Font = new System.Drawing.Font("Segoe UI", 11.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnEsci.ForeColor = System.Drawing.Color.White;
            this.btnEsci.Location = new System.Drawing.Point(574, 9);
            this.btnEsci.Name = "btnEsci";
            this.btnEsci.Size = new System.Drawing.Size(245, 46);
            this.btnEsci.TabIndex = 29;
            this.btnEsci.Text = "❌ Annulla (Esc)";
            this.btnEsci.UseVisualStyleBackColor = false;
            this.btnEsci.Click += new System.EventHandler(this.esciToolStripMenuItem_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Segoe UI", 10.2F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(51)))), ((int)(((byte)(65)))), ((int)(((byte)(85)))));
            this.label1.Location = new System.Drawing.Point(14, 96);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(115, 19);
            this.label1.TabIndex = 2;
            this.label1.Text = "Codice articolo:";
            // 
            // txtArtCod
            // 
            this.txtArtCod.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(241)))), ((int)(((byte)(245)))), ((int)(((byte)(249)))));
            this.txtArtCod.Font = new System.Drawing.Font("Segoe UI", 11.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtArtCod.Location = new System.Drawing.Point(140, 92);
            this.txtArtCod.Name = "txtArtCod";
            this.txtArtCod.ReadOnly = true;
            this.txtArtCod.Size = new System.Drawing.Size(120, 27);
            this.txtArtCod.TabIndex = 3;
            this.txtArtCod.Text = "NEW";
            this.txtArtCod.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Font = new System.Drawing.Font("Segoe UI", 10.2F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label6.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(51)))), ((int)(((byte)(65)))), ((int)(((byte)(85)))));
            this.label6.Location = new System.Drawing.Point(14, 134);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(100, 19);
            this.label6.TabIndex = 9;
            this.label6.Text = "Barcode EAN:";
            // 
            // txtEanEan
            // 
            this.txtEanEan.Font = new System.Drawing.Font("Segoe UI", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtEanEan.Location = new System.Drawing.Point(140, 130);
            this.txtEanEan.MaxLength = 14;
            this.txtEanEan.Name = "txtEanEan";
            this.txtEanEan.Size = new System.Drawing.Size(220, 29);
            this.txtEanEan.TabIndex = 4;
            this.txtEanEan.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtArt_KeyDown);
            this.txtEanEan.Validated += new System.EventHandler(this.txtArt_Validated);
            this.txtEanEan.TextChanged += new System.EventHandler(this.txtEanEan_TextChanged);
            // 
            // btnCercaWeb
            // 
            this.btnCercaWeb.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(37)))), ((int)(((byte)(99)))), ((int)(((byte)(235)))));
            this.btnCercaWeb.Cursor = System.Windows.Forms.Cursors.Hand;
            this.btnCercaWeb.FlatAppearance.BorderSize = 0;
            this.btnCercaWeb.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnCercaWeb.Font = new System.Drawing.Font("Segoe UI", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnCercaWeb.ForeColor = System.Drawing.Color.White;
            this.btnCercaWeb.Location = new System.Drawing.Point(370, 128);
            this.btnCercaWeb.Name = "btnCercaWeb";
            this.btnCercaWeb.Size = new System.Drawing.Size(160, 32);
            this.btnCercaWeb.TabIndex = 30;
            this.btnCercaWeb.Text = "🔍 Cerca Web (F2)";
            this.btnCercaWeb.UseVisualStyleBackColor = false;
            this.btnCercaWeb.Click += new System.EventHandler(this.btnCercaWeb_Click);
            // 
            // lblWebStatus
            // 
            this.lblWebStatus.AutoSize = true;
            this.lblWebStatus.Font = new System.Drawing.Font("Segoe UI", 9.5F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblWebStatus.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(37)))), ((int)(((byte)(99)))), ((int)(((byte)(235)))));
            this.lblWebStatus.Location = new System.Drawing.Point(140, 163);
            this.lblWebStatus.Name = "lblWebStatus";
            this.lblWebStatus.Size = new System.Drawing.Size(262, 17);
            this.lblWebStatus.TabIndex = 31;
            this.lblWebStatus.Text = "💡 Inserisci il Barcode e premi Cerca Web";
            // 
            // picWebProduct
            // 
            this.picWebProduct.BackColor = System.Drawing.Color.White;
            this.picWebProduct.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.picWebProduct.Location = new System.Drawing.Point(670, 130);
            this.picWebProduct.Name = "picWebProduct";
            this.picWebProduct.Size = new System.Drawing.Size(150, 150);
            this.picWebProduct.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom;
            this.picWebProduct.TabIndex = 32;
            this.picWebProduct.TabStop = false;
            // 
            // lblPicCaption
            // 
            this.lblPicCaption.Font = new System.Drawing.Font("Segoe UI", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblPicCaption.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(100)))), ((int)(((byte)(116)))), ((int)(((byte)(139)))));
            this.lblPicCaption.Location = new System.Drawing.Point(670, 283);
            this.lblPicCaption.Name = "lblPicCaption";
            this.lblPicCaption.Size = new System.Drawing.Size(150, 20);
            this.lblPicCaption.TabIndex = 33;
            this.lblPicCaption.Text = "Immagine Prodotto";
            this.lblPicCaption.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Segoe UI", 10.2F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label2.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(51)))), ((int)(((byte)(65)))), ((int)(((byte)(85)))));
            this.label2.Location = new System.Drawing.Point(14, 192);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(92, 19);
            this.label2.TabIndex = 4;
            this.label2.Text = "Descrizione:";
            // 
            // txtArtDes
            // 
            this.txtArtDes.Font = new System.Drawing.Font("Segoe UI", 11.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtArtDes.Location = new System.Drawing.Point(140, 188);
            this.txtArtDes.Name = "txtArtDes";
            this.txtArtDes.Size = new System.Drawing.Size(510, 27);
            this.txtArtDes.TabIndex = 5;
            this.txtArtDes.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtArt_KeyDown);
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Font = new System.Drawing.Font("Segoe UI", 10.2F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label5.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(51)))), ((int)(((byte)(65)))), ((int)(((byte)(85)))));
            this.label5.Location = new System.Drawing.Point(14, 232);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(37, 19);
            this.label5.TabIndex = 8;
            this.label5.Text = "IVA:";
            // 
            // cmbArtIva
            // 
            this.cmbArtIva.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmbArtIva.Font = new System.Drawing.Font("Segoe UI", 10.2F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.cmbArtIva.FormattingEnabled = true;
            this.cmbArtIva.Location = new System.Drawing.Point(140, 228);
            this.cmbArtIva.Name = "cmbArtIva";
            this.cmbArtIva.Size = new System.Drawing.Size(240, 27);
            this.cmbArtIva.TabIndex = 15;
            this.cmbArtIva.SelectionChangeCommitted += new System.EventHandler(this.cmbArtIva_SelectionChangeCommitted);
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("Segoe UI", 10.2F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label3.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(51)))), ((int)(((byte)(65)))), ((int)(((byte)(85)))));
            this.label3.Location = new System.Drawing.Point(14, 270);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(37, 19);
            this.label3.TabIndex = 6;
            this.label3.Text = "UM:";
            // 
            // cmbArtUmi
            // 
            this.cmbArtUmi.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmbArtUmi.Font = new System.Drawing.Font("Segoe UI", 10.2F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.cmbArtUmi.FormattingEnabled = true;
            this.cmbArtUmi.Location = new System.Drawing.Point(140, 266);
            this.cmbArtUmi.Name = "cmbArtUmi";
            this.cmbArtUmi.Size = new System.Drawing.Size(240, 27);
            this.cmbArtUmi.TabIndex = 16;
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Font = new System.Drawing.Font("Segoe UI", 10.2F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label4.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(51)))), ((int)(((byte)(65)))), ((int)(((byte)(85)))));
            this.label4.Location = new System.Drawing.Point(14, 308);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(97, 19);
            this.label4.TabIndex = 7;
            this.label4.Text = "Grammatura:";
            // 
            // cmbArtTgr
            // 
            this.cmbArtTgr.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmbArtTgr.Font = new System.Drawing.Font("Segoe UI", 10.2F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.cmbArtTgr.FormattingEnabled = true;
            this.cmbArtTgr.Location = new System.Drawing.Point(140, 304);
            this.cmbArtTgr.Name = "cmbArtTgr";
            this.cmbArtTgr.Size = new System.Drawing.Size(240, 27);
            this.cmbArtTgr.TabIndex = 17;
            // 
            // label9
            // 
            this.label9.AutoSize = true;
            this.label9.Font = new System.Drawing.Font("Segoe UI", 10.2F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label9.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(51)))), ((int)(((byte)(65)))), ((int)(((byte)(85)))));
            this.label9.Location = new System.Drawing.Point(14, 346);
            this.label9.Name = "label9";
            this.label9.Size = new System.Drawing.Size(120, 19);
            this.label9.TabIndex = 12;
            this.label9.Text = "Contenuto/Peso:";
            // 
            // txtArtNet
            // 
            this.txtArtNet.Font = new System.Drawing.Font("Segoe UI", 11.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtArtNet.Location = new System.Drawing.Point(140, 342);
            this.txtArtNet.MaxLength = 8;
            this.txtArtNet.Name = "txtArtNet";
            this.txtArtNet.Size = new System.Drawing.Size(124, 27);
            this.txtArtNet.TabIndex = 18;
            this.txtArtNet.Text = "1";
            this.txtArtNet.TextAlign = System.Windows.Forms.HorizontalAlignment.Right;
            this.txtArtNet.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtArt_KeyDown);
            this.txtArtNet.Validated += new System.EventHandler(this.txtArt_Validated);
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Font = new System.Drawing.Font("Segoe UI", 10.2F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label7.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(51)))), ((int)(((byte)(65)))), ((int)(((byte)(85)))));
            this.label7.Location = new System.Drawing.Point(14, 384);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(98, 19);
            this.label7.TabIndex = 10;
            this.label7.Text = "Merceologia:";
            this.label7.Visible = false;
            // 
            // cmbArtEc3
            // 
            this.cmbArtEc3.Font = new System.Drawing.Font("Segoe UI", 10.2F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.cmbArtEc3.FormattingEnabled = true;
            this.cmbArtEc3.Location = new System.Drawing.Point(140, 380);
            this.cmbArtEc3.Name = "cmbArtEc3";
            this.cmbArtEc3.Size = new System.Drawing.Size(340, 27);
            this.cmbArtEc3.TabIndex = 19;
            this.cmbArtEc3.Visible = false;
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.Font = new System.Drawing.Font("Segoe UI", 10.2F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label8.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(51)))), ((int)(((byte)(65)))), ((int)(((byte)(85)))));
            this.label8.Location = new System.Drawing.Point(14, 384);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(68, 19);
            this.label8.TabIndex = 11;
            this.label8.Text = "Reparto:";
            // 
            // cmbArtRep
            // 
            this.cmbArtRep.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmbArtRep.Font = new System.Drawing.Font("Segoe UI", 10.2F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.cmbArtRep.FormattingEnabled = true;
            this.cmbArtRep.Location = new System.Drawing.Point(140, 380);
            this.cmbArtRep.Name = "cmbArtRep";
            this.cmbArtRep.Size = new System.Drawing.Size(340, 27);
            this.cmbArtRep.TabIndex = 22;
            // 
            // label10
            // 
            this.label10.AutoSize = true;
            this.label10.Font = new System.Drawing.Font("Segoe UI", 10.2F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label10.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(51)))), ((int)(((byte)(65)))), ((int)(((byte)(85)))));
            this.label10.Location = new System.Drawing.Point(14, 422);
            this.label10.Name = "label10";
            this.label10.Size = new System.Drawing.Size(76, 19);
            this.label10.TabIndex = 13;
            this.label10.Text = "Fornitore:";
            // 
            // cmbLiaFor
            // 
            this.cmbLiaFor.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmbLiaFor.Font = new System.Drawing.Font("Segoe UI", 10.2F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.cmbLiaFor.FormattingEnabled = true;
            this.cmbLiaFor.Location = new System.Drawing.Point(140, 418);
            this.cmbLiaFor.Name = "cmbLiaFor";
            this.cmbLiaFor.Size = new System.Drawing.Size(220, 27);
            this.cmbLiaFor.TabIndex = 24;
            // 
            // lblArtArf
            // 
            this.lblArtArf.AutoSize = true;
            this.lblArtArf.Font = new System.Drawing.Font("Segoe UI", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblArtArf.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(51)))), ((int)(((byte)(65)))), ((int)(((byte)(85)))));
            this.lblArtArf.Location = new System.Drawing.Point(368, 422);
            this.lblArtArf.Name = "lblArtArf";
            this.lblArtArf.Size = new System.Drawing.Size(85, 17);
            this.lblArtArf.TabIndex = 37;
            this.lblArtArf.Text = "Art. Forn:";
            // 
            // txtArtArf
            // 
            this.txtArtArf.Font = new System.Drawing.Font("Segoe UI", 10.5F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtArtArf.Location = new System.Drawing.Point(455, 418);
            this.txtArtArf.MaxLength = 20;
            this.txtArtArf.Name = "txtArtArf";
            this.txtArtArf.Size = new System.Drawing.Size(145, 26);
            this.txtArtArf.TabIndex = 25;
            this.txtArtArf.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtArt_KeyDown);
            this.txtArtArf.Validated += new System.EventHandler(this.txtArt_Validated);
            // 
            // lblCosto
            // 
            this.lblCosto.AutoSize = true;
            this.lblCosto.Font = new System.Drawing.Font("Segoe UI", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblCosto.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(51)))), ((int)(((byte)(65)))), ((int)(((byte)(85)))));
            this.lblCosto.Location = new System.Drawing.Point(608, 422);
            this.lblCosto.Name = "lblCosto";
            this.lblCosto.Size = new System.Drawing.Size(63, 17);
            this.lblCosto.TabIndex = 38;
            this.lblCosto.Text = "Costo €:";
            // 
            // txtLiaPrc
            // 
            this.txtLiaPrc.Font = new System.Drawing.Font("Segoe UI", 11.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtLiaPrc.Location = new System.Drawing.Point(675, 418);
            this.txtLiaPrc.MaxLength = 8;
            this.txtLiaPrc.Name = "txtLiaPrc";
            this.txtLiaPrc.Size = new System.Drawing.Size(145, 27);
            this.txtLiaPrc.TabIndex = 26;
            this.txtLiaPrc.Text = "0";
            this.txtLiaPrc.TextAlign = System.Windows.Forms.HorizontalAlignment.Right;
            this.txtLiaPrc.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtArt_KeyDown);
            this.txtLiaPrc.Validated += new System.EventHandler(this.txtArt_Validated);
            this.txtLiaPrc.TextChanged += new System.EventHandler(this.txtLiaPrc_TextChanged);
            // 
            // label11
            // 
            this.label11.AutoSize = true;
            this.label11.Font = new System.Drawing.Font("Segoe UI", 10.2F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label11.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(51)))), ((int)(((byte)(65)))), ((int)(((byte)(85)))));
            this.label11.Location = new System.Drawing.Point(14, 460);
            this.label11.Name = "label11";
            this.label11.Size = new System.Drawing.Size(61, 19);
            this.label11.TabIndex = 14;
            this.label11.Text = "Listino:";
            // 
            // cmbLivLis
            // 
            this.cmbLivLis.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmbLivLis.Font = new System.Drawing.Font("Segoe UI", 10.2F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.cmbLivLis.FormattingEnabled = true;
            this.cmbLivLis.Location = new System.Drawing.Point(140, 456);
            this.cmbLivLis.Name = "cmbLivLis";
            this.cmbLivLis.Size = new System.Drawing.Size(220, 27);
            this.cmbLivLis.TabIndex = 27;
            // 
            // lblRicarico
            // 
            this.lblRicarico.AutoSize = true;
            this.lblRicarico.Font = new System.Drawing.Font("Segoe UI", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblRicarico.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(51)))), ((int)(((byte)(65)))), ((int)(((byte)(85)))));
            this.lblRicarico.Location = new System.Drawing.Point(368, 460);
            this.lblRicarico.Name = "lblRicarico";
            this.lblRicarico.Size = new System.Drawing.Size(81, 17);
            this.lblRicarico.TabIndex = 39;
            this.lblRicarico.Text = "% Ricarico:";
            // 
            // txtRicarico
            // 
            this.txtRicarico.Font = new System.Drawing.Font("Segoe UI", 10.5F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtRicarico.Location = new System.Drawing.Point(455, 456);
            this.txtRicarico.MaxLength = 6;
            this.txtRicarico.Name = "txtRicarico";
            this.txtRicarico.Size = new System.Drawing.Size(145, 26);
            this.txtRicarico.TabIndex = 28;
            this.txtRicarico.Text = "0";
            this.txtRicarico.TextAlign = System.Windows.Forms.HorizontalAlignment.Right;
            this.txtRicarico.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtArt_KeyDown);
            this.txtRicarico.Validated += new System.EventHandler(this.txtArt_Validated);
            this.txtRicarico.TextChanged += new System.EventHandler(this.txtRicarico_TextChanged);
            // 
            // lblPrezzo
            // 
            this.lblPrezzo.AutoSize = true;
            this.lblPrezzo.Font = new System.Drawing.Font("Segoe UI", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblPrezzo.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(51)))), ((int)(((byte)(65)))), ((int)(((byte)(85)))));
            this.lblPrezzo.Location = new System.Drawing.Point(608, 460);
            this.lblPrezzo.Name = "lblPrezzo";
            this.lblPrezzo.Size = new System.Drawing.Size(68, 17);
            this.lblPrezzo.TabIndex = 40;
            this.lblPrezzo.Text = "Prezzo €:";
            // 
            // txtLivPrv
            // 
            this.txtLivPrv.Font = new System.Drawing.Font("Segoe UI", 11.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtLivPrv.Location = new System.Drawing.Point(675, 456);
            this.txtLivPrv.MaxLength = 8;
            this.txtLivPrv.Name = "txtLivPrv";
            this.txtLivPrv.Size = new System.Drawing.Size(145, 27);
            this.txtLivPrv.TabIndex = 29;
            this.txtLivPrv.Text = "0";
            this.txtLivPrv.TextAlign = System.Windows.Forms.HorizontalAlignment.Right;
            this.txtLivPrv.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtArt_KeyDown);
            this.txtLivPrv.Validated += new System.EventHandler(this.txtArt_Validated);
            // 
            // chkMem
            // 
            this.chkMem.AutoSize = true;
            this.chkMem.Font = new System.Drawing.Font("Segoe UI", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.chkMem.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(71)))), ((int)(((byte)(85)))), ((int)(((byte)(105)))));
            this.chkMem.Location = new System.Drawing.Point(455, 498);
            this.chkMem.Name = "chkMem";
            this.chkMem.Size = new System.Drawing.Size(240, 21);
            this.chkMem.TabIndex = 30;
            this.chkMem.Text = "Proponi ultimi dati inserimenti";
            this.chkMem.UseVisualStyleBackColor = true;
            // 
            // chkRicercaAuto
            // 
            this.chkRicercaAuto.AutoSize = true;
            this.chkRicercaAuto.Checked = true;
            this.chkRicercaAuto.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chkRicercaAuto.Font = new System.Drawing.Font("Segoe UI", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.chkRicercaAuto.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(71)))), ((int)(((byte)(85)))), ((int)(((byte)(105)))));
            this.chkRicercaAuto.Location = new System.Drawing.Point(370, 95);
            this.chkRicercaAuto.Name = "chkRicercaAuto";
            this.chkRicercaAuto.Size = new System.Drawing.Size(235, 21);
            this.chkRicercaAuto.TabIndex = 31;
            this.chkRicercaAuto.Text = "Ricerca automatica da barcode";
            this.chkRicercaAuto.UseVisualStyleBackColor = true;
            this.chkRicercaAuto.CheckedChanged += new System.EventHandler(this.chkRicercaAuto_CheckedChanged);
            // 
            // frmAnaArtNuovo
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(248)))), ((int)(((byte)(250)))), ((int)(((byte)(252)))));
            this.ClientSize = new System.Drawing.Size(840, 615);
            this.ControlBox = false;
            this.Controls.Add(this.panelFooter);
            this.Controls.Add(this.panelHeader);
            this.Controls.Add(this.lblPicCaption);
            this.Controls.Add(this.picWebProduct);
            this.Controls.Add(this.lblWebStatus);
            this.Controls.Add(this.btnCercaWeb);
            this.Controls.Add(this.chkRicercaAuto);
            this.Controls.Add(this.chkMem);
            this.Controls.Add(this.lblPrezzo);
            this.Controls.Add(this.txtLivPrv);
            this.Controls.Add(this.lblRicarico);
            this.Controls.Add(this.txtRicarico);
            this.Controls.Add(this.cmbLivLis);
            this.Controls.Add(this.label11);
            this.Controls.Add(this.lblCosto);
            this.Controls.Add(this.txtLiaPrc);
            this.Controls.Add(this.lblArtArf);
            this.Controls.Add(this.txtArtArf);
            this.Controls.Add(this.cmbLiaFor);
            this.Controls.Add(this.label10);
            this.Controls.Add(this.txtEanEan);
            this.Controls.Add(this.cmbArtRep);
            this.Controls.Add(this.cmbArtEc3);
            this.Controls.Add(this.txtArtNet);
            this.Controls.Add(this.cmbArtTgr);
            this.Controls.Add(this.cmbArtUmi);
            this.Controls.Add(this.cmbArtIva);
            this.Controls.Add(this.label9);
            this.Controls.Add(this.label8);
            this.Controls.Add(this.label7);
            this.Controls.Add(this.label6);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.txtArtDes);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.txtArtCod);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.menuStrip1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.KeyPreview = true;
            this.MainMenuStrip = this.menuStrip1;
            this.Name = "frmAnaArtNuovo";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Inserimento articolo veloce";
            this.Load += new System.EventHandler(this.frmAnaArtNuovo_Load);
            this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmAnaArtNuovo_KeyDown);
            this.panelHeader.ResumeLayout(false);
            this.panelHeader.PerformLayout();
            this.panelFooter.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.picWebProduct)).EndInit();
            this.menuStrip1.ResumeLayout(false);
            this.menuStrip1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.MenuStrip menuStrip1;
        private System.Windows.Forms.ToolStripMenuItem esciToolStripMenuItem;
        private System.Windows.Forms.Panel panelHeader;
        private System.Windows.Forms.Label lblHeaderTitle;
        private System.Windows.Forms.Label lblHeaderSub;
        private System.Windows.Forms.Panel panelFooter;
        private System.Windows.Forms.Button btnSalvaNuovo;
        private System.Windows.Forms.Button btnOk;
        private System.Windows.Forms.Button btnEsci;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox txtArtCod;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox txtArtDes;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.Label label9;
        private System.Windows.Forms.Label label10;
        private System.Windows.Forms.Label label11;
        private System.Windows.Forms.ComboBox cmbArtIva;
        private System.Windows.Forms.ComboBox cmbArtUmi;
        private System.Windows.Forms.ComboBox cmbArtTgr;
        private System.Windows.Forms.TextBox txtArtNet;
        private System.Windows.Forms.ComboBox cmbArtEc3;
        private System.Windows.Forms.ComboBox cmbArtRep;
        private System.Windows.Forms.TextBox txtEanEan;
        private System.Windows.Forms.TextBox txtLiaPrc;
        private System.Windows.Forms.TextBox txtLivPrv;
        private System.Windows.Forms.ComboBox cmbLiaFor;
        private System.Windows.Forms.ComboBox cmbLivLis;
        private System.Windows.Forms.CheckBox chkMem;
        private System.Windows.Forms.CheckBox chkRicercaAuto;
        private System.Windows.Forms.Button btnCercaWeb;
        private System.Windows.Forms.Label lblWebStatus;
        private System.Windows.Forms.PictureBox picWebProduct;
        private System.Windows.Forms.Label lblPicCaption;
        private System.Windows.Forms.Label lblArtArf;
        private System.Windows.Forms.TextBox txtArtArf;
        private System.Windows.Forms.Label lblCosto;
        private System.Windows.Forms.Label lblRicarico;
        private System.Windows.Forms.TextBox txtRicarico;
        private System.Windows.Forms.Label lblPrezzo;
    }
}