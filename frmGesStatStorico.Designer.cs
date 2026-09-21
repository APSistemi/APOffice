namespace APOffice
{
    partial class frmGesStatStorico
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
            this.pdfXRepartoToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.utlitàToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.dgv1 = new APOffice.APDataGridView();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.chkTipMon = new System.Windows.Forms.CheckBox();
            this.dtpDti = new System.Windows.Forms.DateTimePicker();
            this.dtpDtf = new System.Windows.Forms.DateTimePicker();
            this.btnOk = new System.Windows.Forms.Button();
            this.label3 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.progressBar1 = new System.Windows.Forms.ProgressBar();
            this.label6 = new System.Windows.Forms.Label();
            this.cmbNeg = new System.Windows.Forms.ComboBox();
            this.lblVenVal = new System.Windows.Forms.Label();
            this.lblScn = new System.Windows.Forms.Label();
            this.label5 = new System.Windows.Forms.Label();
            this.label13 = new System.Windows.Forms.Label();
            this.lblAcqQta = new System.Windows.Forms.Label();
            this.label15 = new System.Windows.Forms.Label();
            this.label14 = new System.Windows.Forms.Label();
            this.lblAcqVal = new System.Windows.Forms.Label();
            this.panel1 = new System.Windows.Forms.Panel();
            this.label25 = new System.Windows.Forms.Label();
            this.lblVenMrg = new System.Windows.Forms.Label();
            this.lblVenUti = new System.Windows.Forms.Label();
            this.label23 = new System.Windows.Forms.Label();
            this.lblVenVcs = new System.Windows.Forms.Label();
            this.label22 = new System.Windows.Forms.Label();
            this.lblVenVni = new System.Windows.Forms.Label();
            this.label17 = new System.Windows.Forms.Label();
            this.lblVenQkg = new System.Windows.Forms.Label();
            this.lblAcqQkg = new System.Windows.Forms.Label();
            this.label21 = new System.Windows.Forms.Label();
            this.label20 = new System.Windows.Forms.Label();
            this.label19 = new System.Windows.Forms.Label();
            this.label18 = new System.Windows.Forms.Label();
            this.lblVenQta = new System.Windows.Forms.Label();
            this.label16 = new System.Windows.Forms.Label();
            this.chkTipArt = new System.Windows.Forms.CheckBox();
            this.label7 = new System.Windows.Forms.Label();
            this.cancellaStoricoToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.menuStrip1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgv1)).BeginInit();
            this.panel1.SuspendLayout();
            this.SuspendLayout();
            // 
            // menuStrip1
            // 
            this.menuStrip1.BackColor = System.Drawing.SystemColors.ControlLight;
            this.menuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.esciToolStripMenuItem,
            this.pdfXRepartoToolStripMenuItem,
            this.utlitàToolStripMenuItem});
            this.menuStrip1.Location = new System.Drawing.Point(0, 0);
            this.menuStrip1.Name = "menuStrip1";
            this.menuStrip1.Size = new System.Drawing.Size(998, 24);
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
            // pdfXRepartoToolStripMenuItem
            // 
            this.pdfXRepartoToolStripMenuItem.Name = "pdfXRepartoToolStripMenuItem";
            this.pdfXRepartoToolStripMenuItem.Size = new System.Drawing.Size(86, 20);
            this.pdfXRepartoToolStripMenuItem.Text = "Pdf x reparto";
            this.pdfXRepartoToolStripMenuItem.Click += new System.EventHandler(this.pdfXRepartoToolStripMenuItem_Click);
            // 
            // utlitàToolStripMenuItem
            // 
            this.utlitàToolStripMenuItem.Alignment = System.Windows.Forms.ToolStripItemAlignment.Right;
            this.utlitàToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.cancellaStoricoToolStripMenuItem});
            this.utlitàToolStripMenuItem.Name = "utlitàToolStripMenuItem";
            this.utlitàToolStripMenuItem.Size = new System.Drawing.Size(47, 20);
            this.utlitàToolStripMenuItem.Text = "Utlità";
            this.utlitàToolStripMenuItem.Click += new System.EventHandler(this.utlitàToolStripMenuItem_Click);
            // 
            // dgv1
            // 
            this.dgv1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgv1.Location = new System.Drawing.Point(11, 124);
            this.dgv1.Name = "dgv1";
            this.dgv1.Size = new System.Drawing.Size(977, 459);
            this.dgv1.TabIndex = 1;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(309, 64);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(91, 13);
            this.label1.TabIndex = 2;
            this.label1.Text = "Selezione periodo";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(27, 45);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(66, 13);
            this.label2.TabIndex = 3;
            this.label2.Text = "Tipo periodo";
            // 
            // chkTipMon
            // 
            this.chkTipMon.Appearance = System.Windows.Forms.Appearance.Button;
            this.chkTipMon.Font = new System.Drawing.Font("Microsoft Sans Serif", 15.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.chkTipMon.Location = new System.Drawing.Point(117, 31);
            this.chkTipMon.Name = "chkTipMon";
            this.chkTipMon.Size = new System.Drawing.Size(176, 42);
            this.chkTipMon.TabIndex = 4;
            this.chkTipMon.Text = "Mensile";
            this.chkTipMon.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            this.chkTipMon.UseVisualStyleBackColor = true;
            this.chkTipMon.CheckedChanged += new System.EventHandler(this.chkTipMon_CheckedChanged);
            // 
            // dtpDti
            // 
            this.dtpDti.Location = new System.Drawing.Point(438, 43);
            this.dtpDti.Name = "dtpDti";
            this.dtpDti.Size = new System.Drawing.Size(200, 20);
            this.dtpDti.TabIndex = 5;
            // 
            // dtpDtf
            // 
            this.dtpDtf.Location = new System.Drawing.Point(438, 78);
            this.dtpDtf.Name = "dtpDtf";
            this.dtpDtf.Size = new System.Drawing.Size(200, 20);
            this.dtpDtf.TabIndex = 6;
            // 
            // btnOk
            // 
            this.btnOk.Font = new System.Drawing.Font("Microsoft Sans Serif", 18F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnOk.Location = new System.Drawing.Point(705, 66);
            this.btnOk.Name = "btnOk";
            this.btnOk.Size = new System.Drawing.Size(155, 39);
            this.btnOk.TabIndex = 7;
            this.btnOk.Text = "Estrai";
            this.btnOk.UseVisualStyleBackColor = true;
            this.btnOk.Click += new System.EventHandler(this.btnOk_Click);
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(411, 47);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(21, 13);
            this.label3.TabIndex = 8;
            this.label3.Text = "dal";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(415, 83);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(15, 13);
            this.label4.TabIndex = 9;
            this.label4.Text = "al";
            // 
            // progressBar1
            // 
            this.progressBar1.Location = new System.Drawing.Point(0, 677);
            this.progressBar1.Name = "progressBar1";
            this.progressBar1.Size = new System.Drawing.Size(1016, 10);
            this.progressBar1.TabIndex = 10;
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(658, 40);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(46, 13);
            this.label6.TabIndex = 25;
            this.label6.Text = "Negozio";
            // 
            // cmbNeg
            // 
            this.cmbNeg.FormattingEnabled = true;
            this.cmbNeg.Location = new System.Drawing.Point(705, 36);
            this.cmbNeg.Name = "cmbNeg";
            this.cmbNeg.Size = new System.Drawing.Size(155, 21);
            this.cmbNeg.TabIndex = 24;
            // 
            // lblVenVal
            // 
            this.lblVenVal.BackColor = System.Drawing.Color.White;
            this.lblVenVal.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.lblVenVal.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblVenVal.Location = new System.Drawing.Point(383, 43);
            this.lblVenVal.Name = "lblVenVal";
            this.lblVenVal.Size = new System.Drawing.Size(100, 18);
            this.lblVenVal.TabIndex = 37;
            this.lblVenVal.Text = "0";
            this.lblVenVal.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // lblScn
            // 
            this.lblScn.BackColor = System.Drawing.Color.White;
            this.lblScn.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.lblScn.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblScn.Location = new System.Drawing.Point(907, 10);
            this.lblScn.Name = "lblScn";
            this.lblScn.Size = new System.Drawing.Size(61, 18);
            this.lblScn.TabIndex = 35;
            this.lblScn.Text = "0";
            this.lblScn.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(857, 15);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(48, 13);
            this.label5.TabIndex = 34;
            this.label5.Text = "Scontrini";
            // 
            // label13
            // 
            this.label13.AutoSize = true;
            this.label13.Location = new System.Drawing.Point(9, 15);
            this.label13.Name = "label13";
            this.label13.Size = new System.Drawing.Size(57, 13);
            this.label13.TabIndex = 46;
            this.label13.Text = "Acquistato";
            // 
            // lblAcqQta
            // 
            this.lblAcqQta.BackColor = System.Drawing.Color.White;
            this.lblAcqQta.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.lblAcqQta.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblAcqQta.Location = new System.Drawing.Point(111, 13);
            this.lblAcqQta.Name = "lblAcqQta";
            this.lblAcqQta.Size = new System.Drawing.Size(84, 18);
            this.lblAcqQta.TabIndex = 47;
            this.lblAcqQta.Text = "0";
            this.lblAcqQta.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // label15
            // 
            this.label15.AutoSize = true;
            this.label15.Location = new System.Drawing.Point(66, 16);
            this.label15.Name = "label15";
            this.label15.Size = new System.Drawing.Size(42, 13);
            this.label15.TabIndex = 48;
            this.label15.Text = "Q.tà Pz";
            // 
            // label14
            // 
            this.label14.AutoSize = true;
            this.label14.Location = new System.Drawing.Point(344, 22);
            this.label14.Name = "label14";
            this.label14.Size = new System.Drawing.Size(37, 13);
            this.label14.TabIndex = 49;
            this.label14.Text = "Valore";
            // 
            // lblAcqVal
            // 
            this.lblAcqVal.BackColor = System.Drawing.Color.White;
            this.lblAcqVal.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.lblAcqVal.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblAcqVal.Location = new System.Drawing.Point(383, 14);
            this.lblAcqVal.Name = "lblAcqVal";
            this.lblAcqVal.Size = new System.Drawing.Size(100, 18);
            this.lblAcqVal.TabIndex = 50;
            this.lblAcqVal.Text = "0";
            this.lblAcqVal.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // panel1
            // 
            this.panel1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.panel1.Controls.Add(this.label25);
            this.panel1.Controls.Add(this.lblVenMrg);
            this.panel1.Controls.Add(this.lblVenUti);
            this.panel1.Controls.Add(this.label23);
            this.panel1.Controls.Add(this.lblVenVcs);
            this.panel1.Controls.Add(this.label22);
            this.panel1.Controls.Add(this.lblVenVni);
            this.panel1.Controls.Add(this.label17);
            this.panel1.Controls.Add(this.lblVenQkg);
            this.panel1.Controls.Add(this.lblAcqQkg);
            this.panel1.Controls.Add(this.label21);
            this.panel1.Controls.Add(this.label20);
            this.panel1.Controls.Add(this.label19);
            this.panel1.Controls.Add(this.label18);
            this.panel1.Controls.Add(this.lblVenQta);
            this.panel1.Controls.Add(this.label16);
            this.panel1.Controls.Add(this.lblAcqVal);
            this.panel1.Controls.Add(this.label13);
            this.panel1.Controls.Add(this.lblScn);
            this.panel1.Controls.Add(this.label5);
            this.panel1.Controls.Add(this.label14);
            this.panel1.Controls.Add(this.lblAcqQta);
            this.panel1.Controls.Add(this.label15);
            this.panel1.Controls.Add(this.lblVenVal);
            this.panel1.Location = new System.Drawing.Point(11, 589);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(975, 78);
            this.panel1.TabIndex = 51;
            // 
            // label25
            // 
            this.label25.BackColor = System.Drawing.Color.White;
            this.label25.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label25.Location = new System.Drawing.Point(947, 42);
            this.label25.Name = "label25";
            this.label25.Size = new System.Drawing.Size(26, 18);
            this.label25.TabIndex = 66;
            this.label25.Text = "%";
            this.label25.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // lblVenMrg
            // 
            this.lblVenMrg.BackColor = System.Drawing.Color.White;
            this.lblVenMrg.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.lblVenMrg.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblVenMrg.Location = new System.Drawing.Point(907, 43);
            this.lblVenMrg.Name = "lblVenMrg";
            this.lblVenMrg.Size = new System.Drawing.Size(41, 18);
            this.lblVenMrg.TabIndex = 65;
            this.lblVenMrg.Text = "0";
            this.lblVenMrg.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // lblVenUti
            // 
            this.lblVenUti.BackColor = System.Drawing.Color.White;
            this.lblVenUti.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.lblVenUti.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblVenUti.Location = new System.Drawing.Point(828, 43);
            this.lblVenUti.Name = "lblVenUti";
            this.lblVenUti.Size = new System.Drawing.Size(73, 18);
            this.lblVenUti.TabIndex = 64;
            this.lblVenUti.Text = "0";
            this.lblVenUti.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // label23
            // 
            this.label23.AutoSize = true;
            this.label23.Location = new System.Drawing.Point(794, 48);
            this.label23.Name = "label23";
            this.label23.Size = new System.Drawing.Size(28, 13);
            this.label23.TabIndex = 63;
            this.label23.Text = "Utile";
            // 
            // lblVenVcs
            // 
            this.lblVenVcs.BackColor = System.Drawing.Color.White;
            this.lblVenVcs.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.lblVenVcs.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblVenVcs.Location = new System.Drawing.Point(697, 43);
            this.lblVenVcs.Name = "lblVenVcs";
            this.lblVenVcs.Size = new System.Drawing.Size(96, 18);
            this.lblVenVcs.TabIndex = 62;
            this.lblVenVcs.Text = "0";
            this.lblVenVcs.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // label22
            // 
            this.label22.AutoSize = true;
            this.label22.Location = new System.Drawing.Point(663, 48);
            this.label22.Name = "label22";
            this.label22.Size = new System.Drawing.Size(34, 13);
            this.label22.TabIndex = 61;
            this.label22.Text = "Costo";
            // 
            // lblVenVni
            // 
            this.lblVenVni.BackColor = System.Drawing.Color.White;
            this.lblVenVni.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.lblVenVni.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblVenVni.Location = new System.Drawing.Point(563, 43);
            this.lblVenVni.Name = "lblVenVni";
            this.lblVenVni.Size = new System.Drawing.Size(97, 18);
            this.lblVenVni.TabIndex = 60;
            this.lblVenVni.Text = "0";
            this.lblVenVni.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // label17
            // 
            this.label17.AutoSize = true;
            this.label17.Location = new System.Drawing.Point(491, 47);
            this.label17.Name = "label17";
            this.label17.Size = new System.Drawing.Size(68, 13);
            this.label17.TabIndex = 59;
            this.label17.Text = "Valore n/IVA";
            // 
            // lblVenQkg
            // 
            this.lblVenQkg.BackColor = System.Drawing.Color.White;
            this.lblVenQkg.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.lblVenQkg.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblVenQkg.Location = new System.Drawing.Point(254, 45);
            this.lblVenQkg.Name = "lblVenQkg";
            this.lblVenQkg.Size = new System.Drawing.Size(79, 18);
            this.lblVenQkg.TabIndex = 58;
            this.lblVenQkg.Text = "0";
            this.lblVenQkg.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // lblAcqQkg
            // 
            this.lblAcqQkg.BackColor = System.Drawing.Color.White;
            this.lblAcqQkg.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.lblAcqQkg.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblAcqQkg.Location = new System.Drawing.Point(254, 16);
            this.lblAcqQkg.Name = "lblAcqQkg";
            this.lblAcqQkg.Size = new System.Drawing.Size(79, 18);
            this.lblAcqQkg.TabIndex = 57;
            this.lblAcqQkg.Text = "0";
            this.lblAcqQkg.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // label21
            // 
            this.label21.AutoSize = true;
            this.label21.Location = new System.Drawing.Point(206, 45);
            this.label21.Name = "label21";
            this.label21.Size = new System.Drawing.Size(43, 13);
            this.label21.TabIndex = 56;
            this.label21.Text = "Q.tà Kg";
            // 
            // label20
            // 
            this.label20.AutoSize = true;
            this.label20.Location = new System.Drawing.Point(205, 19);
            this.label20.Name = "label20";
            this.label20.Size = new System.Drawing.Size(43, 13);
            this.label20.TabIndex = 55;
            this.label20.Text = "Q.tà Kg";
            // 
            // label19
            // 
            this.label19.AutoSize = true;
            this.label19.Location = new System.Drawing.Point(66, 48);
            this.label19.Name = "label19";
            this.label19.Size = new System.Drawing.Size(42, 13);
            this.label19.TabIndex = 54;
            this.label19.Text = "Q.tà Pz";
            // 
            // label18
            // 
            this.label18.AutoSize = true;
            this.label18.Location = new System.Drawing.Point(343, 48);
            this.label18.Name = "label18";
            this.label18.Size = new System.Drawing.Size(37, 13);
            this.label18.TabIndex = 53;
            this.label18.Text = "Valore";
            // 
            // lblVenQta
            // 
            this.lblVenQta.BackColor = System.Drawing.Color.White;
            this.lblVenQta.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.lblVenQta.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblVenQta.Location = new System.Drawing.Point(111, 43);
            this.lblVenQta.Name = "lblVenQta";
            this.lblVenQta.Size = new System.Drawing.Size(84, 18);
            this.lblVenQta.TabIndex = 52;
            this.lblVenQta.Text = "0";
            this.lblVenQta.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // label16
            // 
            this.label16.AutoSize = true;
            this.label16.Location = new System.Drawing.Point(9, 46);
            this.label16.Name = "label16";
            this.label16.Size = new System.Drawing.Size(47, 13);
            this.label16.TabIndex = 51;
            this.label16.Text = "Venduto";
            // 
            // chkTipArt
            // 
            this.chkTipArt.Appearance = System.Windows.Forms.Appearance.Button;
            this.chkTipArt.Checked = true;
            this.chkTipArt.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chkTipArt.Font = new System.Drawing.Font("Microsoft Sans Serif", 15.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.chkTipArt.Location = new System.Drawing.Point(117, 79);
            this.chkTipArt.Name = "chkTipArt";
            this.chkTipArt.Size = new System.Drawing.Size(176, 42);
            this.chkTipArt.TabIndex = 52;
            this.chkTipArt.Text = "Articolo";
            this.chkTipArt.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            this.chkTipArt.UseVisualStyleBackColor = true;
            this.chkTipArt.CheckedChanged += new System.EventHandler(this.chkTipArt_CheckedChanged);
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Location = new System.Drawing.Point(22, 92);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(75, 13);
            this.label7.TabIndex = 53;
            this.label7.Text = "Tipo selezione";
            // 
            // cancellaStoricoToolStripMenuItem
            // 
            this.cancellaStoricoToolStripMenuItem.Name = "cancellaStoricoToolStripMenuItem";
            this.cancellaStoricoToolStripMenuItem.Size = new System.Drawing.Size(158, 22);
            this.cancellaStoricoToolStripMenuItem.Text = "Cancella storico";
            this.cancellaStoricoToolStripMenuItem.Click += new System.EventHandler(this.cancellaStoricoToolStripMenuItem_Click);
            // 
            // frmGesStatStorico
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(998, 689);
            this.Controls.Add(this.label7);
            this.Controls.Add(this.chkTipArt);
            this.Controls.Add(this.panel1);
            this.Controls.Add(this.cmbNeg);
            this.Controls.Add(this.label6);
            this.Controls.Add(this.progressBar1);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.btnOk);
            this.Controls.Add(this.dtpDtf);
            this.Controls.Add(this.dtpDti);
            this.Controls.Add(this.chkTipMon);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.dgv1);
            this.Controls.Add(this.menuStrip1);
            this.KeyPreview = true;
            this.MainMenuStrip = this.menuStrip1;
            this.Name = "frmGesStatStorico";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Statistiche Storico";
            this.Load += new System.EventHandler(this.frmGesStatStorico_Load);
            this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmGesStatStorico_KeyDown);
            this.menuStrip1.ResumeLayout(false);
            this.menuStrip1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgv1)).EndInit();
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.MenuStrip menuStrip1;
        private System.Windows.Forms.ToolStripMenuItem esciToolStripMenuItem;
        private APOffice.APDataGridView dgv1;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.CheckBox chkTipMon;
        private System.Windows.Forms.DateTimePicker dtpDti;
        private System.Windows.Forms.DateTimePicker dtpDtf;
        private System.Windows.Forms.Button btnOk;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.ProgressBar progressBar1;
        private System.Windows.Forms.ToolStripMenuItem pdfXRepartoToolStripMenuItem;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.ComboBox cmbNeg;
        private System.Windows.Forms.Label lblVenVal;
        private System.Windows.Forms.Label lblScn;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Label label13;
        private System.Windows.Forms.Label lblAcqQta;
        private System.Windows.Forms.Label label15;
        private System.Windows.Forms.Label label14;
        private System.Windows.Forms.Label lblAcqVal;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.Label lblVenQkg;
        private System.Windows.Forms.Label lblAcqQkg;
        private System.Windows.Forms.Label label21;
        private System.Windows.Forms.Label label20;
        private System.Windows.Forms.Label label19;
        private System.Windows.Forms.Label label18;
        private System.Windows.Forms.Label lblVenQta;
        private System.Windows.Forms.Label label16;
        private System.Windows.Forms.Label lblVenVcs;
        private System.Windows.Forms.Label label22;
        private System.Windows.Forms.Label lblVenVni;
        private System.Windows.Forms.Label label17;
        private System.Windows.Forms.Label lblVenUti;
        private System.Windows.Forms.Label label23;
        private System.Windows.Forms.Label label25;
        private System.Windows.Forms.Label lblVenMrg;
        private System.Windows.Forms.CheckBox chkTipArt;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.ToolStripMenuItem utlitàToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem cancellaStoricoToolStripMenuItem;
    }
}
