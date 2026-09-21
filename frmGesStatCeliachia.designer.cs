namespace APOffice
{
    partial class frmGesStatCeliachia
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
            this.excelToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.dgv1 = new APOffice.APDataGridView();
            this.btnEstrai = new System.Windows.Forms.Button();
            this.label6 = new System.Windows.Forms.Label();
            this.cmbNeg = new System.Windows.Forms.ComboBox();
            this.progressBar1 = new System.Windows.Forms.ProgressBar();
            this.lblTot = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.btnInvio = new System.Windows.Forms.Button();
            this.btnDoc = new System.Windows.Forms.Button();
            this.label7 = new System.Windows.Forms.Label();
            this.label8 = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.nudMon = new System.Windows.Forms.NumericUpDown();
            this.cmbYea = new System.Windows.Forms.ComboBox();
            this.lblMon = new System.Windows.Forms.Label();
            this.dtpDdt = new System.Windows.Forms.DateTimePicker();
            this.cmbNdo = new System.Windows.Forms.ComboBox();
            this.label2 = new System.Windows.Forms.Label();
            this.lblNdo = new System.Windows.Forms.Label();
            this.dgv2 = new APOffice.APDataGridView();
            this.lblTotTot = new System.Windows.Forms.Label();
            this.lblTotIva = new System.Windows.Forms.Label();
            this.lblTotImp = new System.Windows.Forms.Label();
            this.label17 = new System.Windows.Forms.Label();
            this.label16 = new System.Windows.Forms.Label();
            this.label15 = new System.Windows.Forms.Label();
            this.lblTotQta = new System.Windows.Forms.Label();
            this.label22 = new System.Windows.Forms.Label();
            this.menuStrip1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgv1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.nudMon)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.dgv2)).BeginInit();
            this.SuspendLayout();
            // 
            // menuStrip1
            // 
            this.menuStrip1.BackColor = System.Drawing.SystemColors.ControlLight;
            this.menuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.esciToolStripMenuItem,
            this.excelToolStripMenuItem});
            this.menuStrip1.Location = new System.Drawing.Point(0, 0);
            this.menuStrip1.Name = "menuStrip1";
            this.menuStrip1.Size = new System.Drawing.Size(904, 24);
            this.menuStrip1.TabIndex = 2;
            this.menuStrip1.Text = "menuStrip1";
            // 
            // esciToolStripMenuItem
            // 
            this.esciToolStripMenuItem.Name = "esciToolStripMenuItem";
            this.esciToolStripMenuItem.Size = new System.Drawing.Size(39, 20);
            this.esciToolStripMenuItem.Text = "Esci";
            this.esciToolStripMenuItem.Click += new System.EventHandler(this.esciToolStripMenuItem_Click);
            // 
            // excelToolStripMenuItem
            // 
            this.excelToolStripMenuItem.Name = "excelToolStripMenuItem";
            this.excelToolStripMenuItem.Size = new System.Drawing.Size(95, 20);
            this.excelToolStripMenuItem.Text = "Excel riepilogo";
            this.excelToolStripMenuItem.Click += new System.EventHandler(this.excelToolStripMenuItem_Click);
            // 
            // dgv1
            // 
            this.dgv1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgv1.Location = new System.Drawing.Point(12, 117);
            this.dgv1.Name = "dgv1";
            this.dgv1.RowHeadersWidth = 20;
            this.dgv1.Size = new System.Drawing.Size(880, 357);
            this.dgv1.TabIndex = 13;
            // 
            // btnEstrai
            // 
            this.btnEstrai.BackColor = System.Drawing.SystemColors.Control;
            this.btnEstrai.Location = new System.Drawing.Point(705, 61);
            this.btnEstrai.Name = "btnEstrai";
            this.btnEstrai.Size = new System.Drawing.Size(186, 38);
            this.btnEstrai.TabIndex = 18;
            this.btnEstrai.Text = "Estrai";
            this.btnEstrai.UseVisualStyleBackColor = false;
            this.btnEstrai.Click += new System.EventHandler(this.btnEstrai_Click);
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(645, 37);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(46, 13);
            this.label6.TabIndex = 25;
            this.label6.Text = "Negozio";
            // 
            // cmbNeg
            // 
            this.cmbNeg.FormattingEnabled = true;
            this.cmbNeg.Location = new System.Drawing.Point(705, 32);
            this.cmbNeg.Name = "cmbNeg";
            this.cmbNeg.Size = new System.Drawing.Size(186, 21);
            this.cmbNeg.TabIndex = 24;
            // 
            // progressBar1
            // 
            this.progressBar1.Location = new System.Drawing.Point(0, 614);
            this.progressBar1.Name = "progressBar1";
            this.progressBar1.Size = new System.Drawing.Size(905, 10);
            this.progressBar1.TabIndex = 26;
            // 
            // lblTot
            // 
            this.lblTot.BackColor = System.Drawing.Color.White;
            this.lblTot.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.lblTot.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblTot.Location = new System.Drawing.Point(591, 490);
            this.lblTot.Name = "lblTot";
            this.lblTot.Size = new System.Drawing.Size(100, 18);
            this.lblTot.TabIndex = 32;
            this.lblTot.Text = "0";
            this.lblTot.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(464, 495);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(37, 13);
            this.label4.TabIndex = 28;
            this.label4.Text = "Totale";
            // 
            // btnInvio
            // 
            this.btnInvio.BackColor = System.Drawing.SystemColors.Control;
            this.btnInvio.Location = new System.Drawing.Point(706, 571);
            this.btnInvio.Name = "btnInvio";
            this.btnInvio.Size = new System.Drawing.Size(186, 27);
            this.btnInvio.TabIndex = 33;
            this.btnInvio.Text = "Invio rendicontazione";
            this.btnInvio.UseVisualStyleBackColor = false;
            this.btnInvio.Click += new System.EventHandler(this.btnInvio_Click);
            // 
            // btnDoc
            // 
            this.btnDoc.BackColor = System.Drawing.SystemColors.Control;
            this.btnDoc.Location = new System.Drawing.Point(706, 535);
            this.btnDoc.Name = "btnDoc";
            this.btnDoc.Size = new System.Drawing.Size(186, 30);
            this.btnDoc.TabIndex = 34;
            this.btnDoc.Text = "Generazione numero documento";
            this.btnDoc.UseVisualStyleBackColor = false;
            this.btnDoc.Click += new System.EventHandler(this.btnDoc_Click);
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Location = new System.Drawing.Point(10, 83);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(62, 13);
            this.label7.TabIndex = 36;
            this.label7.Text = "Documento";
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.Location = new System.Drawing.Point(9, 45);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(32, 13);
            this.label8.TabIndex = 37;
            this.label8.Text = "Anno";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(190, 45);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(33, 13);
            this.label1.TabIndex = 39;
            this.label1.Text = "Mese";
            // 
            // nudMon
            // 
            this.nudMon.Font = new System.Drawing.Font("Microsoft Sans Serif", 14.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.nudMon.Location = new System.Drawing.Point(230, 37);
            this.nudMon.Maximum = new decimal(new int[] {
            12,
            0,
            0,
            0});
            this.nudMon.Minimum = new decimal(new int[] {
            1,
            0,
            0,
            0});
            this.nudMon.Name = "nudMon";
            this.nudMon.Size = new System.Drawing.Size(61, 29);
            this.nudMon.TabIndex = 41;
            this.nudMon.Value = new decimal(new int[] {
            1,
            0,
            0,
            0});
            this.nudMon.ValueChanged += new System.EventHandler(this.nudMon_ValueChanged);
            // 
            // cmbYea
            // 
            this.cmbYea.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmbYea.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.cmbYea.FormattingEnabled = true;
            this.cmbYea.Location = new System.Drawing.Point(78, 37);
            this.cmbYea.Name = "cmbYea";
            this.cmbYea.Size = new System.Drawing.Size(100, 24);
            this.cmbYea.TabIndex = 65;
            this.cmbYea.SelectionChangeCommitted += new System.EventHandler(this.cmbYea_SelectionChangeCommitted);
            // 
            // lblMon
            // 
            this.lblMon.AutoSize = true;
            this.lblMon.Location = new System.Drawing.Point(306, 46);
            this.lblMon.Name = "lblMon";
            this.lblMon.Size = new System.Drawing.Size(16, 13);
            this.lblMon.TabIndex = 66;
            this.lblMon.Text = "...";
            // 
            // dtpDdt
            // 
            this.dtpDdt.Location = new System.Drawing.Point(442, 77);
            this.dtpDdt.Name = "dtpDdt";
            this.dtpDdt.Size = new System.Drawing.Size(200, 20);
            this.dtpDdt.TabIndex = 67;
            // 
            // cmbNdo
            // 
            this.cmbNdo.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmbNdo.FormattingEnabled = true;
            this.cmbNdo.Location = new System.Drawing.Point(78, 76);
            this.cmbNdo.Name = "cmbNdo";
            this.cmbNdo.Size = new System.Drawing.Size(214, 21);
            this.cmbNdo.TabIndex = 68;
            this.cmbNdo.SelectedIndexChanged += new System.EventHandler(this.cmbNdo_SelectedIndexChanged);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(409, 82);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(30, 13);
            this.label2.TabIndex = 69;
            this.label2.Text = "Data";
            // 
            // lblNdo
            // 
            this.lblNdo.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.lblNdo.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblNdo.Location = new System.Drawing.Point(309, 77);
            this.lblNdo.Name = "lblNdo";
            this.lblNdo.Size = new System.Drawing.Size(80, 22);
            this.lblNdo.TabIndex = 70;
            this.lblNdo.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // dgv2
            // 
            this.dgv2.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgv2.Location = new System.Drawing.Point(12, 480);
            this.dgv2.Name = "dgv2";
            this.dgv2.Size = new System.Drawing.Size(446, 118);
            this.dgv2.TabIndex = 71;
            // 
            // lblTotTot
            // 
            this.lblTotTot.BackColor = System.Drawing.SystemColors.ControlLight;
            this.lblTotTot.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.lblTotTot.Location = new System.Drawing.Point(591, 571);
            this.lblTotTot.Name = "lblTotTot";
            this.lblTotTot.Size = new System.Drawing.Size(100, 23);
            this.lblTotTot.TabIndex = 77;
            this.lblTotTot.Text = "0";
            this.lblTotTot.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // lblTotIva
            // 
            this.lblTotIva.BackColor = System.Drawing.SystemColors.ControlLight;
            this.lblTotIva.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.lblTotIva.Location = new System.Drawing.Point(591, 545);
            this.lblTotIva.Name = "lblTotIva";
            this.lblTotIva.Size = new System.Drawing.Size(100, 23);
            this.lblTotIva.TabIndex = 76;
            this.lblTotIva.Text = "0";
            this.lblTotIva.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // lblTotImp
            // 
            this.lblTotImp.BackColor = System.Drawing.SystemColors.ControlLight;
            this.lblTotImp.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.lblTotImp.Location = new System.Drawing.Point(591, 518);
            this.lblTotImp.Name = "lblTotImp";
            this.lblTotImp.Size = new System.Drawing.Size(100, 23);
            this.lblTotImp.TabIndex = 75;
            this.lblTotImp.Text = "0";
            this.lblTotImp.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // label17
            // 
            this.label17.AutoSize = true;
            this.label17.Location = new System.Drawing.Point(465, 555);
            this.label17.Name = "label17";
            this.label17.Size = new System.Drawing.Size(44, 13);
            this.label17.TabIndex = 74;
            this.label17.Text = "Imposta";
            // 
            // label16
            // 
            this.label16.AutoSize = true;
            this.label16.Location = new System.Drawing.Point(465, 528);
            this.label16.Name = "label16";
            this.label16.Size = new System.Drawing.Size(54, 13);
            this.label16.TabIndex = 73;
            this.label16.Text = "Imponibile";
            // 
            // label15
            // 
            this.label15.AutoSize = true;
            this.label15.Location = new System.Drawing.Point(465, 583);
            this.label15.Name = "label15";
            this.label15.Size = new System.Drawing.Size(93, 13);
            this.label15.TabIndex = 72;
            this.label15.Text = "Totale documento";
            // 
            // lblTotQta
            // 
            this.lblTotQta.BackColor = System.Drawing.SystemColors.ControlLight;
            this.lblTotQta.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.lblTotQta.Location = new System.Drawing.Point(805, 495);
            this.lblTotQta.Name = "lblTotQta";
            this.lblTotQta.Size = new System.Drawing.Size(86, 23);
            this.lblTotQta.TabIndex = 79;
            this.lblTotQta.Text = "0";
            this.lblTotQta.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // label22
            // 
            this.label22.AutoSize = true;
            this.label22.Location = new System.Drawing.Point(711, 502);
            this.label22.Name = "label22";
            this.label22.Size = new System.Drawing.Size(32, 13);
            this.label22.TabIndex = 78;
            this.label22.Text = "Pezzi";
            // 
            // frmGesStatCeliachia
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(904, 624);
            this.ControlBox = false;
            this.Controls.Add(this.lblTotQta);
            this.Controls.Add(this.label22);
            this.Controls.Add(this.lblTotTot);
            this.Controls.Add(this.lblTotIva);
            this.Controls.Add(this.lblTotImp);
            this.Controls.Add(this.label17);
            this.Controls.Add(this.label16);
            this.Controls.Add(this.label15);
            this.Controls.Add(this.dgv2);
            this.Controls.Add(this.lblNdo);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.cmbNdo);
            this.Controls.Add(this.dtpDdt);
            this.Controls.Add(this.lblMon);
            this.Controls.Add(this.cmbYea);
            this.Controls.Add(this.nudMon);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.label8);
            this.Controls.Add(this.label7);
            this.Controls.Add(this.btnDoc);
            this.Controls.Add(this.btnInvio);
            this.Controls.Add(this.lblTot);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.progressBar1);
            this.Controls.Add(this.label6);
            this.Controls.Add(this.cmbNeg);
            this.Controls.Add(this.btnEstrai);
            this.Controls.Add(this.dgv1);
            this.Controls.Add(this.menuStrip1);
            this.KeyPreview = true;
            this.Name = "frmGesStatCeliachia";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Statistiche celiachia";
            this.Load += new System.EventHandler(this.frmGesStaIVA_Load);
            this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmGesStatIVA_KeyDown);
            this.menuStrip1.ResumeLayout(false);
            this.menuStrip1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgv1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.nudMon)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.dgv2)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.MenuStrip menuStrip1;
        private System.Windows.Forms.ToolStripMenuItem esciToolStripMenuItem;
        private APOffice.APDataGridView dgv1;
        private System.Windows.Forms.Button btnEstrai;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.ComboBox cmbNeg;
        private System.Windows.Forms.ProgressBar progressBar1;
        private System.Windows.Forms.Label lblTot;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Button btnInvio;
        private System.Windows.Forms.Button btnDoc;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.NumericUpDown nudMon;
        private System.Windows.Forms.ComboBox cmbYea;
        private System.Windows.Forms.Label lblMon;
        private System.Windows.Forms.DateTimePicker dtpDdt;
        private System.Windows.Forms.ComboBox cmbNdo;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label lblNdo;
        private System.Windows.Forms.ToolStripMenuItem excelToolStripMenuItem;
        private APOffice.APDataGridView dgv2;
        private System.Windows.Forms.Label lblTotTot;
        private System.Windows.Forms.Label lblTotIva;
        private System.Windows.Forms.Label lblTotImp;
        private System.Windows.Forms.Label label17;
        private System.Windows.Forms.Label label16;
        private System.Windows.Forms.Label label15;
        private System.Windows.Forms.Label lblTotQta;
        private System.Windows.Forms.Label label22;
    }
}
