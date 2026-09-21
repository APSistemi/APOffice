namespace APOffice
{
    partial class frmGesInventario
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
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle2 = new System.Windows.Forms.DataGridViewCellStyle();
            this.menuStrip1 = new System.Windows.Forms.MenuStrip();
            this.esciToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.importazioneDaTerminalinoToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.utilitàToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.simulatoreToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.toolStripMenuItem2 = new System.Windows.Forms.ToolStripMenuItem();
            this.importAPPhoneToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.dgv1 = new APOffice.APDataGridView();
            this.dtpIntDay = new System.Windows.Forms.DateTimePicker();
            this.label1 = new System.Windows.Forms.Label();
            this.lblIntNum = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.txtIntDes = new System.Windows.Forms.TextBox();
            this.label9 = new System.Windows.Forms.Label();
            this.lblInvPez = new System.Windows.Forms.Label();
            this.lblInvRig = new System.Windows.Forms.Label();
            this.label6 = new System.Windows.Forms.Label();
            this.lblInvTot = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.btnPrn = new System.Windows.Forms.Button();
            this.btnPrnErr = new System.Windows.Forms.Button();
            this.label4 = new System.Windows.Forms.Label();
            this.btnXls = new System.Windows.Forms.Button();
            this.btnAggCosti = new System.Windows.Forms.Button();
            this.progressBar1 = new System.Windows.Forms.ProgressBar();
            this.label5 = new System.Windows.Forms.Label();
            this.txtSeek = new System.Windows.Forms.TextBox();
            this.label7 = new System.Windows.Forms.Label();
            this.cmbIntTip = new System.Windows.Forms.ComboBox();
            this.btnGia = new System.Windows.Forms.Button();
            this.cmbIntRep = new System.Windows.Forms.ComboBox();
            this.label8 = new System.Windows.Forms.Label();
            this.btnDiff = new System.Windows.Forms.Button();
            this.panel1 = new System.Windows.Forms.Panel();
            this.rdbPrnQtaRil = new System.Windows.Forms.RadioButton();
            this.label10 = new System.Windows.Forms.Label();
            this.rdbPrnQtaGia = new System.Windows.Forms.RadioButton();
            this.panel2 = new System.Windows.Forms.Panel();
            this.rdbPrnValCos = new System.Windows.Forms.RadioButton();
            this.label11 = new System.Windows.Forms.Label();
            this.rdbPrnValPrv = new System.Windows.Forms.RadioButton();
            this.chkGia = new System.Windows.Forms.CheckBox();
            this.label12 = new System.Windows.Forms.Label();
            this.cmbIntNeg = new System.Windows.Forms.ComboBox();
            this.label13 = new System.Windows.Forms.Label();
            this.txtRil = new System.Windows.Forms.TextBox();
            this.menuStrip1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgv1)).BeginInit();
            this.panel1.SuspendLayout();
            this.panel2.SuspendLayout();
            this.SuspendLayout();
            // 
            // menuStrip1
            // 
            this.menuStrip1.BackColor = System.Drawing.SystemColors.ControlLight;
            this.menuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.esciToolStripMenuItem,
            this.importazioneDaTerminalinoToolStripMenuItem,
            this.utilitàToolStripMenuItem});
            this.menuStrip1.Location = new System.Drawing.Point(0, 0);
            this.menuStrip1.Name = "menuStrip1";
            this.menuStrip1.Size = new System.Drawing.Size(1016, 24);
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
            // importazioneDaTerminalinoToolStripMenuItem
            // 
            this.importazioneDaTerminalinoToolStripMenuItem.Name = "importazioneDaTerminalinoToolStripMenuItem";
            this.importazioneDaTerminalinoToolStripMenuItem.Size = new System.Drawing.Size(171, 20);
            this.importazioneDaTerminalinoToolStripMenuItem.Text = "Importazione da Terminalino";
            this.importazioneDaTerminalinoToolStripMenuItem.Click += new System.EventHandler(this.importazioneDaTerminalinoToolStripMenuItem_Click);
            // 
            // utilitàToolStripMenuItem
            // 
            this.utilitàToolStripMenuItem.Alignment = System.Windows.Forms.ToolStripItemAlignment.Right;
            this.utilitàToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.simulatoreToolStripMenuItem,
            this.toolStripMenuItem2,
            this.importAPPhoneToolStripMenuItem});
            this.utilitàToolStripMenuItem.Name = "utilitàToolStripMenuItem";
            this.utilitàToolStripMenuItem.Size = new System.Drawing.Size(50, 20);
            this.utilitàToolStripMenuItem.Text = "Utilità";
            // 
            // simulatoreToolStripMenuItem
            // 
            this.simulatoreToolStripMenuItem.Name = "simulatoreToolStripMenuItem";
            this.simulatoreToolStripMenuItem.Size = new System.Drawing.Size(219, 22);
            this.simulatoreToolStripMenuItem.Text = "Simulatore";
            this.simulatoreToolStripMenuItem.Click += new System.EventHandler(this.simulatoreToolStripMenuItem_Click);
            // 
            // toolStripMenuItem2
            // 
            this.toolStripMenuItem2.Name = "toolStripMenuItem2";
            this.toolStripMenuItem2.Size = new System.Drawing.Size(219, 22);
            this.toolStripMenuItem2.Text = "-----------------------------";
            // 
            // importAPPhoneToolStripMenuItem
            // 
            this.importAPPhoneToolStripMenuItem.Name = "importAPPhoneToolStripMenuItem";
            this.importAPPhoneToolStripMenuItem.Size = new System.Drawing.Size(219, 22);
            this.importAPPhoneToolStripMenuItem.Text = "Import AP Phone";
            this.importAPPhoneToolStripMenuItem.Click += new System.EventHandler(this.importAPPhoneToolStripMenuItem_Click);
            // 
            // dgv1
            // 
            this.dgv1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.dgv1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgv1.Location = new System.Drawing.Point(5, 156);
            this.dgv1.Name = "dgv1";
            this.dgv1.RowHeadersWidth = 20;
            this.dgv1.Size = new System.Drawing.Size(1006, 513);
            this.dgv1.TabIndex = 1;
            this.dgv1.CellContentClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dgv1_CellContentClick);
            this.dgv1.CellDoubleClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dgv1_CellDoubleClick);
            this.dgv1.CellEndEdit += new System.Windows.Forms.DataGridViewCellEventHandler(this.dgv1_CellEndEdit);
            this.dgv1.CellValidating += new System.Windows.Forms.DataGridViewCellValidatingEventHandler(this.dgv1_CellValidating);
            this.dgv1.CurrentCellDirtyStateChanged += new System.EventHandler(this.dgv1_CurrentCellDirtyStateChanged);
            this.dgv1.EditingControlShowing += new System.Windows.Forms.DataGridViewEditingControlShowingEventHandler(this.dgv1_EditingControlShowing);
            // 
            // dtpIntDay
            // 
            this.dtpIntDay.Location = new System.Drawing.Point(204, 85);
            this.dtpIntDay.Name = "dtpIntDay";
            this.dtpIntDay.Size = new System.Drawing.Size(200, 20);
            this.dtpIntDay.TabIndex = 2;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(13, 91);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(44, 13);
            this.label1.TabIndex = 3;
            this.label1.Text = "Numero";
            // 
            // lblIntNum
            // 
            this.lblIntNum.BackColor = System.Drawing.Color.White;
            this.lblIntNum.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.lblIntNum.Location = new System.Drawing.Point(84, 86);
            this.lblIntNum.Name = "lblIntNum";
            this.lblIntNum.Size = new System.Drawing.Size(75, 23);
            this.lblIntNum.TabIndex = 4;
            this.lblIntNum.Text = "Numero";
            this.lblIntNum.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(173, 91);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(21, 13);
            this.label2.TabIndex = 5;
            this.label2.Text = "del";
            // 
            // txtIntDes
            // 
            this.txtIntDes.Location = new System.Drawing.Point(127, 29);
            this.txtIntDes.MaxLength = 50;
            this.txtIntDes.Name = "txtIntDes";
            this.txtIntDes.Size = new System.Drawing.Size(277, 20);
            this.txtIntDes.TabIndex = 6;
            // 
            // label9
            // 
            this.label9.AutoSize = true;
            this.label9.Location = new System.Drawing.Point(655, 648);
            this.label9.Name = "label9";
            this.label9.Size = new System.Drawing.Size(32, 13);
            this.label9.TabIndex = 119;
            this.label9.Text = "Pezzi";
            // 
            // lblInvPez
            // 
            this.lblInvPez.BackColor = System.Drawing.Color.White;
            this.lblInvPez.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.lblInvPez.Location = new System.Drawing.Point(688, 640);
            this.lblInvPez.Name = "lblInvPez";
            this.lblInvPez.Size = new System.Drawing.Size(80, 21);
            this.lblInvPez.TabIndex = 118;
            this.lblInvPez.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // lblInvRig
            // 
            this.lblInvRig.BackColor = System.Drawing.Color.White;
            this.lblInvRig.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.lblInvRig.Location = new System.Drawing.Point(562, 640);
            this.lblInvRig.Name = "lblInvRig";
            this.lblInvRig.Size = new System.Drawing.Size(80, 21);
            this.lblInvRig.TabIndex = 117;
            this.lblInvRig.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(518, 647);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(35, 13);
            this.label6.TabIndex = 116;
            this.label6.Text = "Righe";
            // 
            // lblInvTot
            // 
            this.lblInvTot.BackColor = System.Drawing.Color.White;
            this.lblInvTot.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.lblInvTot.Location = new System.Drawing.Point(839, 639);
            this.lblInvTot.Name = "lblInvTot";
            this.lblInvTot.Size = new System.Drawing.Size(80, 21);
            this.lblInvTot.TabIndex = 115;
            this.lblInvTot.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(774, 648);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(66, 13);
            this.label3.TabIndex = 114;
            this.label3.Text = "Totale costo";
            // 
            // btnPrn
            // 
            this.btnPrn.Location = new System.Drawing.Point(804, 81);
            this.btnPrn.Name = "btnPrn";
            this.btnPrn.Size = new System.Drawing.Size(118, 40);
            this.btnPrn.TabIndex = 120;
            this.btnPrn.Text = "Stampa inventario su PDF";
            this.btnPrn.UseVisualStyleBackColor = true;
            this.btnPrn.Click += new System.EventHandler(this.btnPrn_Click);
            // 
            // btnPrnErr
            // 
            this.btnPrnErr.Location = new System.Drawing.Point(579, 81);
            this.btnPrnErr.Name = "btnPrnErr";
            this.btnPrnErr.Size = new System.Drawing.Size(80, 40);
            this.btnPrnErr.TabIndex = 121;
            this.btnPrnErr.Text = "Stampa errori";
            this.btnPrnErr.UseVisualStyleBackColor = true;
            this.btnPrnErr.Click += new System.EventHandler(this.btnPrnErr_Click);
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(10, 35);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(111, 13);
            this.label4.TabIndex = 122;
            this.label4.Text = "Descrizione inventario";
            // 
            // btnXls
            // 
            this.btnXls.Location = new System.Drawing.Point(923, 80);
            this.btnXls.Name = "btnXls";
            this.btnXls.Size = new System.Drawing.Size(82, 40);
            this.btnXls.TabIndex = 123;
            this.btnXls.Text = "Inventario su excel";
            this.btnXls.UseVisualStyleBackColor = true;
            this.btnXls.Click += new System.EventHandler(this.btnXls_Click);
            // 
            // btnAggCosti
            // 
            this.btnAggCosti.Location = new System.Drawing.Point(417, 82);
            this.btnAggCosti.Name = "btnAggCosti";
            this.btnAggCosti.Size = new System.Drawing.Size(155, 40);
            this.btnAggCosti.TabIndex = 124;
            this.btnAggCosti.Text = "Aggiorna costi";
            this.btnAggCosti.UseVisualStyleBackColor = true;
            this.btnAggCosti.Click += new System.EventHandler(this.btnAggCosti_Click);
            // 
            // progressBar1
            // 
            this.progressBar1.Location = new System.Drawing.Point(0, 669);
            this.progressBar1.Name = "progressBar1";
            this.progressBar1.Size = new System.Drawing.Size(1016, 10);
            this.progressBar1.TabIndex = 125;
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(576, 135);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(120, 13);
            this.label5.TabIndex = 126;
            this.label5.Text = "Ricerca descriz. articolo";
            // 
            // txtSeek
            // 
            this.txtSeek.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtSeek.Location = new System.Drawing.Point(696, 124);
            this.txtSeek.MaxLength = 50;
            this.txtSeek.Name = "txtSeek";
            this.txtSeek.Size = new System.Drawing.Size(310, 26);
            this.txtSeek.TabIndex = 127;
            this.txtSeek.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtSeek_KeyDown);
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Location = new System.Drawing.Point(414, 132);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(28, 13);
            this.label7.TabIndex = 128;
            this.label7.Text = "Tipo";
            // 
            // cmbIntTip
            // 
            this.cmbIntTip.FormattingEnabled = true;
            this.cmbIntTip.Items.AddRange(new object[] {
            "Totale",
            "Parziale",
            "Giacenza"});
            this.cmbIntTip.Location = new System.Drawing.Point(445, 125);
            this.cmbIntTip.Name = "cmbIntTip";
            this.cmbIntTip.Size = new System.Drawing.Size(127, 21);
            this.cmbIntTip.TabIndex = 129;
            // 
            // btnGia
            // 
            this.btnGia.Location = new System.Drawing.Point(417, 27);
            this.btnGia.Name = "btnGia";
            this.btnGia.Size = new System.Drawing.Size(155, 26);
            this.btnGia.TabIndex = 131;
            this.btnGia.Text = "Caricamento giacenza";
            this.btnGia.UseVisualStyleBackColor = true;
            this.btnGia.Click += new System.EventHandler(this.btnGia_Click);
            // 
            // cmbIntRep
            // 
            this.cmbIntRep.FormattingEnabled = true;
            this.cmbIntRep.Location = new System.Drawing.Point(175, 125);
            this.cmbIntRep.Name = "cmbIntRep";
            this.cmbIntRep.Size = new System.Drawing.Size(229, 21);
            this.cmbIntRep.TabIndex = 132;
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.Location = new System.Drawing.Point(80, 130);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(81, 13);
            this.label8.TabIndex = 133;
            this.label8.Text = "Filtro sul reparto";
            // 
            // btnDiff
            // 
            this.btnDiff.Location = new System.Drawing.Point(660, 81);
            this.btnDiff.Name = "btnDiff";
            this.btnDiff.Size = new System.Drawing.Size(142, 40);
            this.btnDiff.TabIndex = 135;
            this.btnDiff.Text = "Stampa differenze tra giacenza e contato";
            this.btnDiff.UseVisualStyleBackColor = true;
            this.btnDiff.Click += new System.EventHandler(this.btnDiff_Click);
            // 
            // panel1
            // 
            this.panel1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.panel1.Controls.Add(this.rdbPrnQtaRil);
            this.panel1.Controls.Add(this.label10);
            this.panel1.Controls.Add(this.rdbPrnQtaGia);
            this.panel1.Location = new System.Drawing.Point(579, 29);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(355, 24);
            this.panel1.TabIndex = 136;
            // 
            // rdbPrnQtaRil
            // 
            this.rdbPrnQtaRil.AutoSize = true;
            this.rdbPrnQtaRil.Checked = true;
            this.rdbPrnQtaRil.Location = new System.Drawing.Point(138, 3);
            this.rdbPrnQtaRil.Name = "rdbPrnQtaRil";
            this.rdbPrnQtaRil.Size = new System.Drawing.Size(59, 17);
            this.rdbPrnQtaRil.TabIndex = 139;
            this.rdbPrnQtaRil.TabStop = true;
            this.rdbPrnQtaRil.Text = "rilevato";
            this.rdbPrnQtaRil.UseVisualStyleBackColor = true;
            this.rdbPrnQtaRil.CheckedChanged += new System.EventHandler(this.rdbPrnQtaRil_CheckedChanged);
            // 
            // label10
            // 
            this.label10.AutoSize = true;
            this.label10.Location = new System.Drawing.Point(8, 4);
            this.label10.Name = "label10";
            this.label10.Size = new System.Drawing.Size(101, 13);
            this.label10.TabIndex = 138;
            this.label10.Text = "Stampa quantità su ";
            // 
            // rdbPrnQtaGia
            // 
            this.rdbPrnQtaGia.AutoSize = true;
            this.rdbPrnQtaGia.Location = new System.Drawing.Point(211, 3);
            this.rdbPrnQtaGia.Name = "rdbPrnQtaGia";
            this.rdbPrnQtaGia.Size = new System.Drawing.Size(68, 17);
            this.rdbPrnQtaGia.TabIndex = 137;
            this.rdbPrnQtaGia.Text = "giacenza";
            this.rdbPrnQtaGia.UseVisualStyleBackColor = true;
            this.rdbPrnQtaGia.CheckedChanged += new System.EventHandler(this.rdbPrnQtaRil_CheckedChanged);
            // 
            // panel2
            // 
            this.panel2.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.panel2.Controls.Add(this.rdbPrnValCos);
            this.panel2.Controls.Add(this.label11);
            this.panel2.Controls.Add(this.rdbPrnValPrv);
            this.panel2.Location = new System.Drawing.Point(579, 53);
            this.panel2.Name = "panel2";
            this.panel2.Size = new System.Drawing.Size(355, 23);
            this.panel2.TabIndex = 141;
            // 
            // rdbPrnValCos
            // 
            this.rdbPrnValCos.AutoSize = true;
            this.rdbPrnValCos.Checked = true;
            this.rdbPrnValCos.Location = new System.Drawing.Point(138, 3);
            this.rdbPrnValCos.Name = "rdbPrnValCos";
            this.rdbPrnValCos.Size = new System.Drawing.Size(51, 17);
            this.rdbPrnValCos.TabIndex = 139;
            this.rdbPrnValCos.TabStop = true;
            this.rdbPrnValCos.Text = "costo";
            this.rdbPrnValCos.UseVisualStyleBackColor = true;
            // 
            // label11
            // 
            this.label11.AutoSize = true;
            this.label11.Location = new System.Drawing.Point(8, 6);
            this.label11.Name = "label11";
            this.label11.Size = new System.Drawing.Size(129, 13);
            this.label11.TabIndex = 138;
            this.label11.Text = "Stampa valorizzazione su ";
            // 
            // rdbPrnValPrv
            // 
            this.rdbPrnValPrv.AutoSize = true;
            this.rdbPrnValPrv.Location = new System.Drawing.Point(211, 2);
            this.rdbPrnValPrv.Name = "rdbPrnValPrv";
            this.rdbPrnValPrv.Size = new System.Drawing.Size(105, 17);
            this.rdbPrnValPrv.TabIndex = 137;
            this.rdbPrnValPrv.Text = "prezzo di vendita";
            this.rdbPrnValPrv.UseVisualStyleBackColor = true;
            // 
            // chkGia
            // 
            this.chkGia.AutoSize = true;
            this.chkGia.Enabled = false;
            this.chkGia.Location = new System.Drawing.Point(417, 58);
            this.chkGia.Name = "chkGia";
            this.chkGia.Size = new System.Drawing.Size(91, 17);
            this.chkGia.TabIndex = 142;
            this.chkGia.Text = "Con giacenza";
            this.chkGia.UseVisualStyleBackColor = true;
            // 
            // label12
            // 
            this.label12.AutoSize = true;
            this.label12.Location = new System.Drawing.Point(10, 63);
            this.label12.Name = "label12";
            this.label12.Size = new System.Drawing.Size(46, 13);
            this.label12.TabIndex = 143;
            this.label12.Text = "Negozio";
            // 
            // cmbIntNeg
            // 
            this.cmbIntNeg.FormattingEnabled = true;
            this.cmbIntNeg.Location = new System.Drawing.Point(127, 53);
            this.cmbIntNeg.Name = "cmbIntNeg";
            this.cmbIntNeg.Size = new System.Drawing.Size(277, 21);
            this.cmbIntNeg.TabIndex = 144;
            // 
            // label13
            // 
            this.label13.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.label13.Location = new System.Drawing.Point(932, 29);
            this.label13.Name = "label13";
            this.label13.Size = new System.Drawing.Size(72, 45);
            this.label13.TabIndex = 145;
            this.label13.Text = "Rilevazione term.ril";
            this.label13.TextAlign = System.Drawing.ContentAlignment.TopCenter;
            // 
            // txtRil
            // 
            this.txtRil.Location = new System.Drawing.Point(933, 57);
            this.txtRil.Name = "txtRil";
            this.txtRil.Size = new System.Drawing.Size(68, 20);
            this.txtRil.TabIndex = 146;
            // 
            // frmGesInventario
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1016, 681);
            this.Controls.Add(this.txtRil);
            this.Controls.Add(this.label13);
            this.Controls.Add(this.cmbIntNeg);
            this.Controls.Add(this.label12);
            this.Controls.Add(this.chkGia);
            this.Controls.Add(this.panel2);
            this.Controls.Add(this.panel1);
            this.Controls.Add(this.btnDiff);
            this.Controls.Add(this.label8);
            this.Controls.Add(this.cmbIntRep);
            this.Controls.Add(this.btnGia);
            this.Controls.Add(this.cmbIntTip);
            this.Controls.Add(this.label7);
            this.Controls.Add(this.txtSeek);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.progressBar1);
            this.Controls.Add(this.btnAggCosti);
            this.Controls.Add(this.btnXls);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.btnPrnErr);
            this.Controls.Add(this.btnPrn);
            this.Controls.Add(this.label9);
            this.Controls.Add(this.lblInvPez);
            this.Controls.Add(this.lblInvRig);
            this.Controls.Add(this.label6);
            this.Controls.Add(this.lblInvTot);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.txtIntDes);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.lblIntNum);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.dtpIntDay);
            this.Controls.Add(this.dgv1);
            this.Controls.Add(this.menuStrip1);
            this.KeyPreview = true;
            this.MainMenuStrip = this.menuStrip1;
            this.MinimumSize = new System.Drawing.Size(960, 560);
            this.Name = "frmGesInventario";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Gestione Inventario";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.frmGesInventario_FormClosing);
            this.Load += new System.EventHandler(this.frmGesInventario_Load);
            this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmGesInventario_KeyDown);
            this.menuStrip1.ResumeLayout(false);
            this.menuStrip1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgv1)).EndInit();
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            this.panel2.ResumeLayout(false);
            this.panel2.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.MenuStrip menuStrip1;
        private System.Windows.Forms.ToolStripMenuItem esciToolStripMenuItem;
        private APOffice.APDataGridView dgv1;
        private System.Windows.Forms.DateTimePicker dtpIntDay;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label lblIntNum;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox txtIntDes;
        private System.Windows.Forms.Label label9;
        private System.Windows.Forms.Label lblInvPez;
        private System.Windows.Forms.Label lblInvRig;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.Label lblInvTot;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.ToolStripMenuItem importazioneDaTerminalinoToolStripMenuItem;
        private System.Windows.Forms.Button btnPrn;
        private System.Windows.Forms.Button btnPrnErr;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Button btnXls;
        private System.Windows.Forms.Button btnAggCosti;
        private System.Windows.Forms.ToolStripMenuItem utilitàToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem simulatoreToolStripMenuItem;
        private System.Windows.Forms.ProgressBar progressBar1;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.TextBox txtSeek;
        private System.Windows.Forms.ToolStripMenuItem toolStripMenuItem2;
        private System.Windows.Forms.ToolStripMenuItem importAPPhoneToolStripMenuItem;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.ComboBox cmbIntTip;
        private System.Windows.Forms.Button btnGia;
        private System.Windows.Forms.ComboBox cmbIntRep;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.Button btnDiff;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.RadioButton rdbPrnQtaRil;
        private System.Windows.Forms.Label label10;
        private System.Windows.Forms.RadioButton rdbPrnQtaGia;
        private System.Windows.Forms.Panel panel2;
        private System.Windows.Forms.RadioButton rdbPrnValCos;
        private System.Windows.Forms.Label label11;
        private System.Windows.Forms.RadioButton rdbPrnValPrv;
        private System.Windows.Forms.CheckBox chkGia;
        private System.Windows.Forms.Label label12;
        private System.Windows.Forms.ComboBox cmbIntNeg;
        private System.Windows.Forms.Label label13;
        private System.Windows.Forms.TextBox txtRil;
    }
}
