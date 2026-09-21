namespace APOffice
{
    partial class frmGesOfferta
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
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle1 = new System.Windows.Forms.DataGridViewCellStyle();
            this.menuStrip1 = new System.Windows.Forms.MenuStrip();
            this.esciToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.invioANegoziToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.importDaTerminalinoToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.stampaToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.dgv1 = new APOffice.APDataGridView();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.chkOffAnn = new System.Windows.Forms.CheckBox();
            this.txtOffSeek = new System.Windows.Forms.TextBox();
            this.label13 = new System.Windows.Forms.Label();
            this.pnlMxN = new System.Windows.Forms.Panel();
            this.lblOfaMix = new System.Windows.Forms.Label();
            this.label11 = new System.Windows.Forms.Label();
            this.label5 = new System.Windows.Forms.Label();
            this.txtOfaXem = new System.Windows.Forms.TextBox();
            this.label4 = new System.Windows.Forms.Label();
            this.txtOfaXen = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.rdb003 = new System.Windows.Forms.RadioButton();
            this.txtOfaVal = new System.Windows.Forms.TextBox();
            this.rdb001 = new System.Windows.Forms.RadioButton();
            this.rdb002 = new System.Windows.Forms.RadioButton();
            this.label1 = new System.Windows.Forms.Label();
            this.txtSeek = new System.Windows.Forms.TextBox();
            this.btnSeek = new System.Windows.Forms.Button();
            this.label6 = new System.Windows.Forms.Label();
            this.dtpOftDti = new System.Windows.Forms.DateTimePicker();
            this.label7 = new System.Windows.Forms.Label();
            this.dtpOftDtf = new System.Windows.Forms.DateTimePicker();
            this.pnlOft = new System.Windows.Forms.Panel();
            this.cmbOftNeg = new System.Windows.Forms.ComboBox();
            this.label14 = new System.Windows.Forms.Label();
            this.cmbOftCam = new System.Windows.Forms.ComboBox();
            this.label12 = new System.Windows.Forms.Label();
            this.lblOftNum = new System.Windows.Forms.Label();
            this.label10 = new System.Windows.Forms.Label();
            this.cmbOftSta = new System.Windows.Forms.ComboBox();
            this.lblOftYea = new System.Windows.Forms.Label();
            this.txtOftDes = new System.Windows.Forms.TextBox();
            this.label9 = new System.Windows.Forms.Label();
            this.lblOftCod = new System.Windows.Forms.Label();
            this.label8 = new System.Windows.Forms.Label();
            this.btnMix = new System.Windows.Forms.Button();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.label15 = new System.Windows.Forms.Label();
            this.rdbEtiNor = new System.Windows.Forms.RadioButton();
            this.rdbEtiOff = new System.Windows.Forms.RadioButton();
            this.btnEtiPrn = new System.Windows.Forms.Button();
            this.progressBar1 = new System.Windows.Forms.ProgressBar();
            this.menuStrip1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgv1)).BeginInit();
            this.groupBox1.SuspendLayout();
            this.pnlMxN.SuspendLayout();
            this.pnlOft.SuspendLayout();
            this.groupBox2.SuspendLayout();
            this.SuspendLayout();
            // 
            // menuStrip1
            // 
            this.menuStrip1.BackColor = System.Drawing.SystemColors.ControlLight;
            this.menuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.esciToolStripMenuItem,
            this.invioANegoziToolStripMenuItem,
            this.importDaTerminalinoToolStripMenuItem,
            this.stampaToolStripMenuItem});
            this.menuStrip1.Location = new System.Drawing.Point(0, 0);
            this.menuStrip1.Name = "menuStrip1";
            this.menuStrip1.Size = new System.Drawing.Size(690, 24);
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
            // invioANegoziToolStripMenuItem
            // 
            this.invioANegoziToolStripMenuItem.Alignment = System.Windows.Forms.ToolStripItemAlignment.Right;
            this.invioANegoziToolStripMenuItem.Name = "invioANegoziToolStripMenuItem";
            this.invioANegoziToolStripMenuItem.Size = new System.Drawing.Size(92, 20);
            this.invioANegoziToolStripMenuItem.Text = "Invio a negozi";
            this.invioANegoziToolStripMenuItem.Click += new System.EventHandler(this.invioANegoziToolStripMenuItem_Click);
            // 
            // importDaTerminalinoToolStripMenuItem
            // 
            this.importDaTerminalinoToolStripMenuItem.Alignment = System.Windows.Forms.ToolStripItemAlignment.Right;
            this.importDaTerminalinoToolStripMenuItem.Name = "importDaTerminalinoToolStripMenuItem";
            this.importDaTerminalinoToolStripMenuItem.Size = new System.Drawing.Size(135, 20);
            this.importDaTerminalinoToolStripMenuItem.Text = "Import da terminalino";
            this.importDaTerminalinoToolStripMenuItem.Click += new System.EventHandler(this.importDaTerminalinoToolStripMenuItem_Click);
            // 
            // stampaToolStripMenuItem
            // 
            this.stampaToolStripMenuItem.Name = "stampaToolStripMenuItem";
            this.stampaToolStripMenuItem.Size = new System.Drawing.Size(45, 20);
            this.stampaToolStripMenuItem.Text = "Excel";
            this.stampaToolStripMenuItem.Click += new System.EventHandler(this.stampaToolStripMenuItem_Click);
            // 
            // dgv1
            // 
            this.dgv1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            dataGridViewCellStyle1.BackColor = System.Drawing.Color.Lavender;
            this.dgv1.AlternatingRowsDefaultCellStyle = dataGridViewCellStyle1;
            this.dgv1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgv1.Location = new System.Drawing.Point(12, 296);
            this.dgv1.Name = "dgv1";
            this.dgv1.RowHeadersWidth = 20;
            this.dgv1.Size = new System.Drawing.Size(756, 328);
            this.dgv1.TabIndex = 1;
            this.dgv1.CellContentClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dgv1_CellContentClick);
            this.dgv1.CellDoubleClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dgv1_CellDoubleClick);
            this.dgv1.CellEndEdit += new System.Windows.Forms.DataGridViewCellEventHandler(this.dgv1_CellEndEdit);
            this.dgv1.CurrentCellChanged += new System.EventHandler(this.dgv1_CurrentCellChanged);
            this.dgv1.CurrentCellDirtyStateChanged += new System.EventHandler(this.dgv1_CurrentCellDirtyStateChanged);
            this.dgv1.EditingControlShowing += new System.Windows.Forms.DataGridViewEditingControlShowingEventHandler(this.dgv1_EditingControlShowing);
            this.dgv1.DoubleClick += new System.EventHandler(this.dgv1_DoubleClick);
            // 
            // groupBox1
            // 
            this.groupBox1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.groupBox1.Controls.Add(this.chkOffAnn);
            this.groupBox1.Controls.Add(this.txtOffSeek);
            this.groupBox1.Controls.Add(this.label13);
            this.groupBox1.Controls.Add(this.pnlMxN);
            this.groupBox1.Controls.Add(this.label2);
            this.groupBox1.Controls.Add(this.rdb003);
            this.groupBox1.Controls.Add(this.txtOfaVal);
            this.groupBox1.Controls.Add(this.rdb001);
            this.groupBox1.Controls.Add(this.rdb002);
            this.groupBox1.Font = new System.Drawing.Font("Segoe UI", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.groupBox1.Location = new System.Drawing.Point(12, 199);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(756, 90);
            this.groupBox1.TabIndex = 2;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Tipo offerta";
            // 
            // chkOffAnn
            // 
            this.chkOffAnn.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.chkOffAnn.AutoSize = true;
            this.chkOffAnn.CheckAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.chkOffAnn.Location = new System.Drawing.Point(676, 68);
            this.chkOffAnn.Name = "chkOffAnn";
            this.chkOffAnn.Size = new System.Drawing.Size(73, 19);
            this.chkOffAnn.TabIndex = 9;
            this.chkOffAnn.Text = "Annullati";
            this.chkOffAnn.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            this.chkOffAnn.UseVisualStyleBackColor = true;
            this.chkOffAnn.CheckedChanged += new System.EventHandler(this.chkOffAnn_CheckedChanged);
            // 
            // txtOffSeek
            // 
            this.txtOffSeek.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.txtOffSeek.Location = new System.Drawing.Point(435, 42);
            this.txtOffSeek.Name = "txtOffSeek";
            this.txtOffSeek.Size = new System.Drawing.Size(314, 23);
            this.txtOffSeek.TabIndex = 8;
            this.txtOffSeek.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtOffSeek_KeyDown);
            // 
            // label13
            // 
            this.label13.AutoSize = true;
            this.label13.Location = new System.Drawing.Point(438, 23);
            this.label13.Name = "label13";
            this.label13.Size = new System.Drawing.Size(107, 15);
            this.label13.TabIndex = 7;
            this.label13.Text = "Ricerca nell\'offerta";
            // 
            // pnlMxN
            // 
            this.pnlMxN.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.pnlMxN.Controls.Add(this.lblOfaMix);
            this.pnlMxN.Controls.Add(this.label11);
            this.pnlMxN.Controls.Add(this.label5);
            this.pnlMxN.Controls.Add(this.txtOfaXem);
            this.pnlMxN.Controls.Add(this.label4);
            this.pnlMxN.Controls.Add(this.txtOfaXen);
            this.pnlMxN.Controls.Add(this.label3);
            this.pnlMxN.Location = new System.Drawing.Point(211, 20);
            this.pnlMxN.Name = "pnlMxN";
            this.pnlMxN.Size = new System.Drawing.Size(212, 61);
            this.pnlMxN.TabIndex = 2;
            // 
            // lblOfaMix
            // 
            this.lblOfaMix.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.lblOfaMix.Location = new System.Drawing.Point(116, 27);
            this.lblOfaMix.Name = "lblOfaMix";
            this.lblOfaMix.Size = new System.Drawing.Size(78, 23);
            this.lblOfaMix.TabIndex = 11;
            this.lblOfaMix.Text = "0";
            this.lblOfaMix.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // label11
            // 
            this.label11.AutoSize = true;
            this.label11.Location = new System.Drawing.Point(113, 5);
            this.label11.Name = "label11";
            this.label11.Size = new System.Drawing.Size(101, 15);
            this.label11.TabIndex = 10;
            this.label11.Text = "Codice Mix Match";
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(72, 5);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(16, 15);
            this.label5.TabIndex = 9;
            this.label5.Text = "N";
            // 
            // txtOfaXem
            // 
            this.txtOfaXem.Location = new System.Drawing.Point(10, 28);
            this.txtOfaXem.Name = "txtOfaXem";
            this.txtOfaXem.Size = new System.Drawing.Size(31, 23);
            this.txtOfaXem.TabIndex = 3;
            this.txtOfaXem.Text = "0";
            this.txtOfaXem.TextAlign = System.Windows.Forms.HorizontalAlignment.Right;
            this.txtOfaXem.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtOfaVal_KeyDown);
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(47, 30);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(13, 15);
            this.label4.TabIndex = 8;
            this.label4.Text = "x";
            // 
            // txtOfaXen
            // 
            this.txtOfaXen.Location = new System.Drawing.Point(65, 28);
            this.txtOfaXen.Name = "txtOfaXen";
            this.txtOfaXen.Size = new System.Drawing.Size(31, 23);
            this.txtOfaXen.TabIndex = 4;
            this.txtOfaXen.Text = "0";
            this.txtOfaXen.TextAlign = System.Windows.Forms.HorizontalAlignment.Right;
            this.txtOfaXen.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtOfaVal_KeyDown);
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(16, 6);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(18, 15);
            this.label3.TabIndex = 7;
            this.label3.Text = "M";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(137, 20);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(39, 15);
            this.label2.TabIndex = 6;
            this.label2.Text = "Valore";
            // 
            // rdb003
            // 
            this.rdb003.AutoSize = true;
            this.rdb003.Location = new System.Drawing.Point(10, 64);
            this.rdb003.Name = "rdb003";
            this.rdb003.Size = new System.Drawing.Size(51, 19);
            this.rdb003.TabIndex = 2;
            this.rdb003.TabStop = true;
            this.rdb003.Text = "MxN";
            this.rdb003.UseVisualStyleBackColor = true;
            this.rdb003.Click += new System.EventHandler(this.rdbOff_Click);
            // 
            // txtOfaVal
            // 
            this.txtOfaVal.Location = new System.Drawing.Point(137, 39);
            this.txtOfaVal.Name = "txtOfaVal";
            this.txtOfaVal.Size = new System.Drawing.Size(52, 23);
            this.txtOfaVal.TabIndex = 1;
            this.txtOfaVal.TextAlign = System.Windows.Forms.HorizontalAlignment.Right;
            this.txtOfaVal.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtOfaVal_KeyDown);
            // 
            // rdb001
            // 
            this.rdb001.AutoSize = true;
            this.rdb001.Checked = true;
            this.rdb001.Location = new System.Drawing.Point(10, 19);
            this.rdb001.Name = "rdb001";
            this.rdb001.Size = new System.Drawing.Size(95, 19);
            this.rdb001.TabIndex = 0;
            this.rdb001.TabStop = true;
            this.rdb001.Text = "Taglio prezzo";
            this.rdb001.UseVisualStyleBackColor = true;
            this.rdb001.Click += new System.EventHandler(this.rdbOff_Click);
            // 
            // rdb002
            // 
            this.rdb002.AutoSize = true;
            this.rdb002.Location = new System.Drawing.Point(10, 42);
            this.rdb002.Name = "rdb002";
            this.rdb002.Size = new System.Drawing.Size(73, 19);
            this.rdb002.TabIndex = 0;
            this.rdb002.TabStop = true;
            this.rdb002.Text = "Sconto %";
            this.rdb002.UseVisualStyleBackColor = true;
            this.rdb002.Click += new System.EventHandler(this.rdbOff_Click);
            // 
            // label1
            // 
            this.label1.Font = new System.Drawing.Font("Segoe UI", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(13, 142);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(71, 38);
            this.label1.TabIndex = 3;
            this.label1.Text = "Ricerca nuovo articolo";
            // 
            // txtSeek
            // 
            this.txtSeek.Font = new System.Drawing.Font("Segoe UI", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtSeek.Location = new System.Drawing.Point(85, 146);
            this.txtSeek.Name = "txtSeek";
            this.txtSeek.Size = new System.Drawing.Size(260, 27);
            this.txtSeek.TabIndex = 5;
            this.txtSeek.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtSeek_KeyDown);
            // 
            // btnSeek
            // 
            this.btnSeek.Location = new System.Drawing.Point(350, 145);
            this.btnSeek.Name = "btnSeek";
            this.btnSeek.Size = new System.Drawing.Size(46, 29);
            this.btnSeek.TabIndex = 5;
            this.btnSeek.Text = "...";
            this.btnSeek.UseVisualStyleBackColor = true;
            this.btnSeek.Click += new System.EventHandler(this.btnSeek_Click);
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(340, 15);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(27, 15);
            this.label6.TabIndex = 6;
            this.label6.Text = "Dal ";
            // 
            // dtpOftDti
            // 
            this.dtpOftDti.Location = new System.Drawing.Point(369, 11);
            this.dtpOftDti.Name = "dtpOftDti";
            this.dtpOftDti.Size = new System.Drawing.Size(165, 23);
            this.dtpOftDti.TabIndex = 7;
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Location = new System.Drawing.Point(540, 15);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(19, 15);
            this.label7.TabIndex = 8;
            this.label7.Text = "al ";
            // 
            // dtpOftDtf
            // 
            this.dtpOftDtf.Location = new System.Drawing.Point(565, 11);
            this.dtpOftDtf.Name = "dtpOftDtf";
            this.dtpOftDtf.Size = new System.Drawing.Size(165, 23);
            this.dtpOftDtf.TabIndex = 9;
            // 
            // pnlOft
            // 
            this.pnlOft.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.pnlOft.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.pnlOft.Controls.Add(this.cmbOftNeg);
            this.pnlOft.Controls.Add(this.label14);
            this.pnlOft.Controls.Add(this.cmbOftCam);
            this.pnlOft.Controls.Add(this.label12);
            this.pnlOft.Controls.Add(this.lblOftNum);
            this.pnlOft.Controls.Add(this.label10);
            this.pnlOft.Controls.Add(this.cmbOftSta);
            this.pnlOft.Controls.Add(this.lblOftYea);
            this.pnlOft.Controls.Add(this.txtOftDes);
            this.pnlOft.Controls.Add(this.label9);
            this.pnlOft.Controls.Add(this.lblOftCod);
            this.pnlOft.Controls.Add(this.label8);
            this.pnlOft.Controls.Add(this.dtpOftDti);
            this.pnlOft.Controls.Add(this.dtpOftDtf);
            this.pnlOft.Controls.Add(this.label6);
            this.pnlOft.Controls.Add(this.label7);
            this.pnlOft.Font = new System.Drawing.Font("Segoe UI", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.pnlOft.Location = new System.Drawing.Point(12, 27);
            this.pnlOft.Name = "pnlOft";
            this.pnlOft.Size = new System.Drawing.Size(756, 102);
            this.pnlOft.TabIndex = 10;
            // 
            // cmbOftNeg
            // 
            this.cmbOftNeg.FormattingEnabled = true;
            this.cmbOftNeg.Location = new System.Drawing.Point(495, 69);
            this.cmbOftNeg.Name = "cmbOftNeg";
            this.cmbOftNeg.Size = new System.Drawing.Size(235, 23);
            this.cmbOftNeg.TabIndex = 21;
            // 
            // label14
            // 
            this.label14.AutoSize = true;
            this.label14.Location = new System.Drawing.Point(438, 73);
            this.label14.Name = "label14";
            this.label14.Size = new System.Drawing.Size(51, 15);
            this.label14.TabIndex = 20;
            this.label14.Text = "Negozio";
            // 
            // cmbOftCam
            // 
            this.cmbOftCam.FormattingEnabled = true;
            this.cmbOftCam.Location = new System.Drawing.Point(120, 69);
            this.cmbOftCam.Name = "cmbOftCam";
            this.cmbOftCam.Size = new System.Drawing.Size(295, 23);
            this.cmbOftCam.TabIndex = 19;
            // 
            // label12
            // 
            this.label12.AutoSize = true;
            this.label12.Location = new System.Drawing.Point(8, 73);
            this.label12.Name = "label12";
            this.label12.Size = new System.Drawing.Size(102, 15);
            this.label12.TabIndex = 18;
            this.label12.Text = "Campagna fidelity";
            // 
            // lblOftNum
            // 
            this.lblOftNum.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.lblOftNum.Font = new System.Drawing.Font("Segoe UI", 9.75F, System.Drawing.FontStyle.Bold);
            this.lblOftNum.Location = new System.Drawing.Point(235, 10);
            this.lblOftNum.Name = "lblOftNum";
            this.lblOftNum.Size = new System.Drawing.Size(85, 25);
            this.lblOftNum.TabIndex = 12;
            this.lblOftNum.Text = "...";
            this.lblOftNum.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // label10
            // 
            this.label10.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.label10.AutoSize = true;
            this.label10.Location = new System.Drawing.Point(545, 43);
            this.label10.Name = "label10";
            this.label10.Size = new System.Drawing.Size(34, 15);
            this.label10.TabIndex = 16;
            this.label10.Text = "Stato";
            // 
            // cmbOftSta
            // 
            this.cmbOftSta.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.cmbOftSta.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmbOftSta.FormattingEnabled = true;
            this.cmbOftSta.Location = new System.Drawing.Point(585, 39);
            this.cmbOftSta.Name = "cmbOftSta";
            this.cmbOftSta.Size = new System.Drawing.Size(145, 23);
            this.cmbOftSta.TabIndex = 15;
            // 
            // lblOftYea
            // 
            this.lblOftYea.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.lblOftYea.Font = new System.Drawing.Font("Segoe UI", 9.75F, System.Drawing.FontStyle.Bold);
            this.lblOftYea.Location = new System.Drawing.Point(120, 10);
            this.lblOftYea.Name = "lblOftYea";
            this.lblOftYea.Size = new System.Drawing.Size(52, 25);
            this.lblOftYea.TabIndex = 14;
            this.lblOftYea.Text = "...";
            this.lblOftYea.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // txtOftDes
            // 
            this.txtOftDes.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.txtOftDes.Location = new System.Drawing.Point(120, 39);
            this.txtOftDes.MaxLength = 50;
            this.txtOftDes.Name = "txtOftDes";
            this.txtOftDes.Size = new System.Drawing.Size(415, 23);
            this.txtOftDes.TabIndex = 13;
            // 
            // label9
            // 
            this.label9.AutoSize = true;
            this.label9.Location = new System.Drawing.Point(8, 43);
            this.label9.Name = "label9";
            this.label9.Size = new System.Drawing.Size(67, 15);
            this.label9.TabIndex = 12;
            this.label9.Text = "Descrizione";
            // 
            // lblOftCod
            // 
            this.lblOftCod.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.lblOftCod.Font = new System.Drawing.Font("Segoe UI", 9.75F, System.Drawing.FontStyle.Bold);
            this.lblOftCod.Location = new System.Drawing.Point(177, 10);
            this.lblOftCod.Name = "lblOftCod";
            this.lblOftCod.Size = new System.Drawing.Size(53, 25);
            this.lblOftCod.TabIndex = 11;
            this.lblOftCod.Text = "...";
            this.lblOftCod.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.Location = new System.Drawing.Point(8, 15);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(112, 15);
            this.label8.TabIndex = 10;
            this.label8.Text = "Codice promozione";
            // 
            // btnMix
            // 
            this.btnMix.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btnMix.Location = new System.Drawing.Point(606, 134);
            this.btnMix.Name = "btnMix";
            this.btnMix.Size = new System.Drawing.Size(162, 59);
            this.btnMix.TabIndex = 11;
            this.btnMix.Text = "Attribuzione mix-match";
            this.btnMix.UseVisualStyleBackColor = true;
            this.btnMix.Click += new System.EventHandler(this.btnMix_Click);
            // 
            // groupBox2
            // 
            this.groupBox2.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.groupBox2.Controls.Add(this.label15);
            this.groupBox2.Controls.Add(this.rdbEtiNor);
            this.groupBox2.Controls.Add(this.rdbEtiOff);
            this.groupBox2.Controls.Add(this.btnEtiPrn);
            this.groupBox2.Font = new System.Drawing.Font("Segoe UI", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.groupBox2.Location = new System.Drawing.Point(408, 134);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(190, 59);
            this.groupBox2.TabIndex = 18;
            this.groupBox2.TabStop = false;
            // 
            // label15
            // 
            this.label15.AutoSize = true;
            this.label15.Location = new System.Drawing.Point(6, 13);
            this.label15.Name = "label15";
            this.label15.Size = new System.Drawing.Size(41, 13);
            this.label15.TabIndex = 19;
            this.label15.Text = "Prezzo";
            // 
            // rdbEtiNor
            // 
            this.rdbEtiNor.AutoSize = true;
            this.rdbEtiNor.Location = new System.Drawing.Point(118, 11);
            this.rdbEtiNor.Name = "rdbEtiNor";
            this.rdbEtiNor.Size = new System.Drawing.Size(66, 17);
            this.rdbEtiNor.TabIndex = 18;
            this.rdbEtiNor.TabStop = true;
            this.rdbEtiNor.Text = "normale";
            this.rdbEtiNor.UseVisualStyleBackColor = true;
            // 
            // rdbEtiOff
            // 
            this.rdbEtiOff.AutoSize = true;
            this.rdbEtiOff.Checked = true;
            this.rdbEtiOff.Location = new System.Drawing.Point(52, 11);
            this.rdbEtiOff.Name = "rdbEtiOff";
            this.rdbEtiOff.Size = new System.Drawing.Size(59, 17);
            this.rdbEtiOff.TabIndex = 7;
            this.rdbEtiOff.TabStop = true;
            this.rdbEtiOff.Text = "offerta";
            this.rdbEtiOff.UseVisualStyleBackColor = true;
            // 
            // btnEtiPrn
            // 
            this.btnEtiPrn.Location = new System.Drawing.Point(6, 29);
            this.btnEtiPrn.Name = "btnEtiPrn";
            this.btnEtiPrn.Size = new System.Drawing.Size(178, 25);
            this.btnEtiPrn.TabIndex = 17;
            this.btnEtiPrn.Text = "Genera etichette ";
            this.btnEtiPrn.UseVisualStyleBackColor = true;
            this.btnEtiPrn.Click += new System.EventHandler(this.btnEtiPrn_Click);
            // 
            // progressBar1
            // 
            this.progressBar1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.progressBar1.Location = new System.Drawing.Point(0, 626);
            this.progressBar1.Name = "progressBar1";
            this.progressBar1.Size = new System.Drawing.Size(780, 6);
            this.progressBar1.TabIndex = 19;
            // 
            // frmGesOfferta
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(780, 634);
            this.Controls.Add(this.progressBar1);
            this.Controls.Add(this.groupBox2);
            this.Controls.Add(this.btnMix);
            this.Controls.Add(this.pnlOft);
            this.Controls.Add(this.btnSeek);
            this.Controls.Add(this.txtSeek);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.dgv1);
            this.Controls.Add(this.menuStrip1);
            this.KeyPreview = true;
            this.MainMenuStrip = this.menuStrip1;
            this.MinimumSize = new System.Drawing.Size(760, 580);
            this.Name = "frmGesOfferta";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Gestione offerta per singolo articolo";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.frmGesOfferta_FormClosing);
            this.Load += new System.EventHandler(this.frmGesOfferta_Load);
            this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmGesOfferta_KeyDown);
            this.menuStrip1.ResumeLayout(false);
            this.menuStrip1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgv1)).EndInit();
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.pnlMxN.ResumeLayout(false);
            this.pnlMxN.PerformLayout();
            this.pnlOft.ResumeLayout(false);
            this.pnlOft.PerformLayout();
            this.groupBox2.ResumeLayout(false);
            this.groupBox2.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.MenuStrip menuStrip1;
        private System.Windows.Forms.ToolStripMenuItem esciToolStripMenuItem;
        private APOffice.APDataGridView dgv1;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.RadioButton rdb002;
        private System.Windows.Forms.TextBox txtOfaVal;
        private System.Windows.Forms.RadioButton rdb001;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox txtSeek;
        private System.Windows.Forms.Button btnSeek;
        private System.Windows.Forms.TextBox txtOfaXem;
        private System.Windows.Forms.RadioButton rdb003;
        private System.Windows.Forms.TextBox txtOfaXen;
        private System.Windows.Forms.Panel pnlMxN;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.DateTimePicker dtpOftDti;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.DateTimePicker dtpOftDtf;
        private System.Windows.Forms.Panel pnlOft;
        private System.Windows.Forms.TextBox txtOftDes;
        private System.Windows.Forms.Label label9;
        private System.Windows.Forms.Label lblOftCod;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.Label lblOftYea;
        private System.Windows.Forms.Label label10;
        private System.Windows.Forms.ComboBox cmbOftSta;
        private System.Windows.Forms.Label lblOfaMix;
        private System.Windows.Forms.Label label11;
        private System.Windows.Forms.Button btnMix;
        private System.Windows.Forms.Label lblOftNum;
        private System.Windows.Forms.ComboBox cmbOftCam;
        private System.Windows.Forms.Label label12;
        private System.Windows.Forms.TextBox txtOffSeek;
        private System.Windows.Forms.Label label13;
        private System.Windows.Forms.ToolStripMenuItem importDaTerminalinoToolStripMenuItem;
        private System.Windows.Forms.ComboBox cmbOftNeg;
        private System.Windows.Forms.Label label14;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.Label label15;
        private System.Windows.Forms.RadioButton rdbEtiNor;
        private System.Windows.Forms.RadioButton rdbEtiOff;
        private System.Windows.Forms.Button btnEtiPrn;
        private System.Windows.Forms.ProgressBar progressBar1;
        private System.Windows.Forms.ToolStripMenuItem stampaToolStripMenuItem;
        private System.Windows.Forms.CheckBox chkOffAnn;
        private System.Windows.Forms.ToolStripMenuItem invioANegoziToolStripMenuItem;
    }
}
