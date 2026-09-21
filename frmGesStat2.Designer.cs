namespace APOffice
{
    partial class frmGesStat2
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
            this.utilityToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.aggiornamentoCostiToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.btnEstrai = new System.Windows.Forms.Button();
            this.dtpFin = new System.Windows.Forms.DateTimePicker();
            this.dtpIni = new System.Windows.Forms.DateTimePicker();
            this.label2 = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.dgv1 = new APOffice.APDataGridView();
            this.lblTotAcq = new System.Windows.Forms.Label();
            this.label5 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.lblTotVen = new System.Windows.Forms.Label();
            this.lblTotMov = new System.Windows.Forms.Label();
            this.label6 = new System.Windows.Forms.Label();
            this.btnArt = new System.Windows.Forms.Button();
            this.button1 = new System.Windows.Forms.Button();
            this.button2 = new System.Windows.Forms.Button();
            this.button3 = new System.Windows.Forms.Button();
            this.button4 = new System.Windows.Forms.Button();
            this.button5 = new System.Windows.Forms.Button();
            this.button6 = new System.Windows.Forms.Button();
            this.button7 = new System.Windows.Forms.Button();
            this.button8 = new System.Windows.Forms.Button();
            this.button9 = new System.Windows.Forms.Button();
            this.button10 = new System.Windows.Forms.Button();
            this.button11 = new System.Windows.Forms.Button();
            this.button12 = new System.Windows.Forms.Button();
            this.button13 = new System.Windows.Forms.Button();
            this.menuStrip1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgv1)).BeginInit();
            this.SuspendLayout();
            // 
            // menuStrip1
            // 
            this.menuStrip1.BackColor = System.Drawing.SystemColors.ControlLight;
            this.menuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.esciToolStripMenuItem,
            this.utilityToolStripMenuItem});
            this.menuStrip1.Location = new System.Drawing.Point(0, 0);
            this.menuStrip1.Name = "menuStrip1";
            this.menuStrip1.Size = new System.Drawing.Size(998, 24);
            this.menuStrip1.TabIndex = 1;
            this.menuStrip1.Text = "menuStrip1";
            // 
            // esciToolStripMenuItem
            // 
            this.esciToolStripMenuItem.Name = "esciToolStripMenuItem";
            this.esciToolStripMenuItem.Size = new System.Drawing.Size(39, 20);
            this.esciToolStripMenuItem.Text = "Esci";
            this.esciToolStripMenuItem.Click += new System.EventHandler(this.esciToolStripMenuItem_Click);
            // 
            // utilityToolStripMenuItem
            // 
            this.utilityToolStripMenuItem.Alignment = System.Windows.Forms.ToolStripItemAlignment.Right;
            this.utilityToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.aggiornamentoCostiToolStripMenuItem});
            this.utilityToolStripMenuItem.Name = "utilityToolStripMenuItem";
            this.utilityToolStripMenuItem.Size = new System.Drawing.Size(50, 20);
            this.utilityToolStripMenuItem.Text = "Utility";
            // 
            // aggiornamentoCostiToolStripMenuItem
            // 
            this.aggiornamentoCostiToolStripMenuItem.Name = "aggiornamentoCostiToolStripMenuItem";
            this.aggiornamentoCostiToolStripMenuItem.Size = new System.Drawing.Size(186, 22);
            this.aggiornamentoCostiToolStripMenuItem.Text = "Aggiornamento costi";
            this.aggiornamentoCostiToolStripMenuItem.Click += new System.EventHandler(this.aggiornamentoCostiToolStripMenuItem_Click);
            // 
            // btnEstrai
            // 
            this.btnEstrai.BackColor = System.Drawing.SystemColors.Control;
            this.btnEstrai.Location = new System.Drawing.Point(490, 31);
            this.btnEstrai.Name = "btnEstrai";
            this.btnEstrai.Size = new System.Drawing.Size(160, 28);
            this.btnEstrai.TabIndex = 11;
            this.btnEstrai.Text = "Estrai";
            this.btnEstrai.UseVisualStyleBackColor = false;
            this.btnEstrai.Click += new System.EventHandler(this.btnEstrai_Click);
            // 
            // dtpFin
            // 
            this.dtpFin.Location = new System.Drawing.Point(278, 35);
            this.dtpFin.Name = "dtpFin";
            this.dtpFin.Size = new System.Drawing.Size(200, 20);
            this.dtpFin.TabIndex = 10;
            // 
            // dtpIni
            // 
            this.dtpIni.Location = new System.Drawing.Point(41, 34);
            this.dtpIni.Name = "dtpIni";
            this.dtpIni.Size = new System.Drawing.Size(200, 20);
            this.dtpIni.TabIndex = 9;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(254, 38);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(15, 13);
            this.label2.TabIndex = 8;
            this.label2.Text = "al";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(12, 38);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(23, 13);
            this.label1.TabIndex = 7;
            this.label1.Text = "Dal";
            // 
            // dgv1
            // 
            this.dgv1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.dgv1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgv1.Location = new System.Drawing.Point(12, 70);
            this.dgv1.Name = "dgv1";
            this.dgv1.RowHeadersWidth = 20;
            this.dgv1.Size = new System.Drawing.Size(735, 555);
            this.dgv1.TabIndex = 12;
            this.dgv1.CellFormatting += new System.Windows.Forms.DataGridViewCellFormattingEventHandler(this.dgv1_CellFormatting);
            this.dgv1.DoubleClick += new System.EventHandler(this.dgv1_DoubleClick);
            // 
            // lblTotAcq
            // 
            this.lblTotAcq.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.lblTotAcq.BackColor = System.Drawing.Color.White;
            this.lblTotAcq.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.lblTotAcq.Location = new System.Drawing.Point(378, 642);
            this.lblTotAcq.Name = "lblTotAcq";
            this.lblTotAcq.Size = new System.Drawing.Size(100, 23);
            this.lblTotAcq.TabIndex = 21;
            this.lblTotAcq.Text = "0";
            this.lblTotAcq.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // label5
            // 
            this.label5.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(301, 647);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(76, 13);
            this.label5.TabIndex = 20;
            this.label5.Text = "Totale acquisti";
            // 
            // label4
            // 
            this.label4.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(496, 647);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(79, 13);
            this.label4.TabIndex = 19;
            this.label4.Text = "Totale venduto";
            // 
            // lblTotVen
            // 
            this.lblTotVen.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.lblTotVen.BackColor = System.Drawing.Color.White;
            this.lblTotVen.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.lblTotVen.Location = new System.Drawing.Point(581, 642);
            this.lblTotVen.Name = "lblTotVen";
            this.lblTotVen.Size = new System.Drawing.Size(100, 23);
            this.lblTotVen.TabIndex = 18;
            this.lblTotVen.Text = "0";
            this.lblTotVen.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // lblTotMov
            // 
            this.lblTotMov.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.lblTotMov.BackColor = System.Drawing.Color.White;
            this.lblTotMov.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.lblTotMov.Location = new System.Drawing.Point(195, 642);
            this.lblTotMov.Name = "lblTotMov";
            this.lblTotMov.Size = new System.Drawing.Size(100, 23);
            this.lblTotMov.TabIndex = 23;
            this.lblTotMov.Text = "0";
            this.lblTotMov.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // label6
            // 
            this.label6.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(110, 647);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(87, 13);
            this.label6.TabIndex = 22;
            this.label6.Text = "Totale movimenti";
            // 
            // btnArt
            // 
            this.btnArt.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btnArt.Location = new System.Drawing.Point(760, 40);
            this.btnArt.Name = "btnArt";
            this.btnArt.Size = new System.Drawing.Size(220, 36);
            this.btnArt.TabIndex = 24;
            this.btnArt.Text = "Dettaglio per articoli";
            this.btnArt.UseVisualStyleBackColor = true;
            this.btnArt.Click += new System.EventHandler(this.btnArt_Click);
            // 
            // button2
            // 
            this.button2.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.button2.Location = new System.Drawing.Point(760, 80);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(220, 36);
            this.button2.TabIndex = 26;
            this.button2.Text = "Dettaglio per reparto";
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Click += new System.EventHandler(this.button2_Click);
            // 
            // button10
            // 
            this.button10.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.button10.Location = new System.Drawing.Point(760, 120);
            this.button10.Name = "button10";
            this.button10.Size = new System.Drawing.Size(220, 36);
            this.button10.TabIndex = 34;
            this.button10.Text = "Dettaglio per merceologia";
            this.button10.UseVisualStyleBackColor = true;
            this.button10.Click += new System.EventHandler(this.button10_Click);
            // 
            // button12
            // 
            this.button12.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.button12.Location = new System.Drawing.Point(760, 160);
            this.button12.Name = "button12";
            this.button12.Size = new System.Drawing.Size(220, 36);
            this.button12.TabIndex = 36;
            this.button12.Text = "Estrazioni settimanali";
            this.button12.UseVisualStyleBackColor = true;
            this.button12.Click += new System.EventHandler(this.button12_Click);
            // 
            // button9
            // 
            this.button9.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.button9.Location = new System.Drawing.Point(760, 200);
            this.button9.Name = "button9";
            this.button9.Size = new System.Drawing.Size(220, 36);
            this.button9.TabIndex = 33;
            this.button9.Text = "Dettaglio per IVA";
            this.button9.UseVisualStyleBackColor = true;
            this.button9.Click += new System.EventHandler(this.button9_Click);
            // 
            // button1
            // 
            this.button1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.button1.Location = new System.Drawing.Point(760, 240);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(220, 36);
            this.button1.TabIndex = 25;
            this.button1.Text = "Dettaglio per scontrino";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // button3
            // 
            this.button3.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.button3.Location = new System.Drawing.Point(760, 280);
            this.button3.Name = "button3";
            this.button3.Size = new System.Drawing.Size(220, 36);
            this.button3.TabIndex = 27;
            this.button3.Text = "Premi fidelity";
            this.button3.UseVisualStyleBackColor = true;
            this.button3.Click += new System.EventHandler(this.button3_Click);
            // 
            // button7
            // 
            this.button7.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.button7.Location = new System.Drawing.Point(760, 320);
            this.button7.Name = "button7";
            this.button7.Size = new System.Drawing.Size(220, 36);
            this.button7.TabIndex = 31;
            this.button7.Text = "Movimenti punti fidelity";
            this.button7.UseVisualStyleBackColor = true;
            this.button7.Click += new System.EventHandler(this.button7_Click);
            // 
            // button4
            // 
            this.button4.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.button4.Location = new System.Drawing.Point(760, 360);
            this.button4.Name = "button4";
            this.button4.Size = new System.Drawing.Size(220, 36);
            this.button4.TabIndex = 28;
            this.button4.Text = "Movimenti cassa";
            this.button4.UseVisualStyleBackColor = true;
            this.button4.Click += new System.EventHandler(this.button4_Click);
            // 
            // button8
            // 
            this.button8.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.button8.Location = new System.Drawing.Point(760, 400);
            this.button8.Name = "button8";
            this.button8.Size = new System.Drawing.Size(220, 36);
            this.button8.TabIndex = 32;
            this.button8.Text = "Movimenti cassa 2";
            this.button8.UseVisualStyleBackColor = true;
            this.button8.Click += new System.EventHandler(this.button8_Click);
            // 
            // button5
            // 
            this.button5.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.button5.Location = new System.Drawing.Point(760, 440);
            this.button5.Name = "button5";
            this.button5.Size = new System.Drawing.Size(220, 36);
            this.button5.TabIndex = 29;
            this.button5.Text = "Rotazione per articoli";
            this.button5.UseVisualStyleBackColor = true;
            this.button5.Click += new System.EventHandler(this.button5_Click);
            // 
            // button6
            // 
            this.button6.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.button6.Location = new System.Drawing.Point(760, 480);
            this.button6.Name = "button6";
            this.button6.Size = new System.Drawing.Size(220, 36);
            this.button6.TabIndex = 30;
            this.button6.Text = "Movimentazione per articolo";
            this.button6.UseVisualStyleBackColor = true;
            this.button6.Click += new System.EventHandler(this.button6_Click);
            // 
            // button13
            // 
            this.button13.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.button13.Location = new System.Drawing.Point(760, 520);
            this.button13.Name = "button13";
            this.button13.Size = new System.Drawing.Size(220, 36);
            this.button13.TabIndex = 37;
            this.button13.Text = "Celiachia";
            this.button13.UseVisualStyleBackColor = true;
            this.button13.Click += new System.EventHandler(this.button13_Click);
            // 
            // button11
            // 
            this.button11.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.button11.Location = new System.Drawing.Point(760, 560);
            this.button11.Name = "button11";
            this.button11.Size = new System.Drawing.Size(220, 36);
            this.button11.TabIndex = 35;
            this.button11.Text = "Storico";
            this.button11.UseVisualStyleBackColor = true;
            this.button11.Click += new System.EventHandler(this.button11_Click);
            // 
            // frmGesStat2
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1000, 680);
            this.MinimumSize = new System.Drawing.Size(960, 660);
            this.ControlBox = true;
            this.MaximizeBox = true;
            this.MinimizeBox = true;
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Sizable;
            this.Controls.Add(this.button11);
            this.Controls.Add(this.button13);
            this.Controls.Add(this.button6);
            this.Controls.Add(this.button5);
            this.Controls.Add(this.button8);
            this.Controls.Add(this.button4);
            this.Controls.Add(this.button7);
            this.Controls.Add(this.button3);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.button9);
            this.Controls.Add(this.button12);
            this.Controls.Add(this.button10);
            this.Controls.Add(this.button2);
            this.Controls.Add(this.btnArt);
            this.Controls.Add(this.lblTotMov);
            this.Controls.Add(this.label6);
            this.Controls.Add(this.lblTotAcq);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.lblTotVen);
            this.Controls.Add(this.dgv1);
            this.Controls.Add(this.btnEstrai);
            this.Controls.Add(this.dtpFin);
            this.Controls.Add(this.dtpIni);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.menuStrip1);
            this.KeyPreview = true;
            this.Name = "frmGesStat2";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Statistiche";
            this.Load += new System.EventHandler(this.frmGesStat2_Load);
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.frmGesStat2_FormClosing);
            this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmGesStat2_KeyDown);
            this.menuStrip1.ResumeLayout(false);
            this.menuStrip1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgv1)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();
        }

        #endregion

        private System.Windows.Forms.MenuStrip menuStrip1;
        private System.Windows.Forms.ToolStripMenuItem esciToolStripMenuItem;
        private System.Windows.Forms.Button btnEstrai;
        private System.Windows.Forms.DateTimePicker dtpFin;
        private System.Windows.Forms.DateTimePicker dtpIni;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label1;
        private APOffice.APDataGridView dgv1;
        private System.Windows.Forms.Label lblTotAcq;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Label lblTotVen;
        private System.Windows.Forms.Label lblTotMov;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.Button btnArt;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.Button button2;
        private System.Windows.Forms.Button button3;
        private System.Windows.Forms.Button button4;
        private System.Windows.Forms.ToolStripMenuItem utilityToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem aggiornamentoCostiToolStripMenuItem;
        private System.Windows.Forms.Button button5;
        private System.Windows.Forms.Button button6;
        private System.Windows.Forms.Button button7;
        private System.Windows.Forms.Button button8;
        private System.Windows.Forms.Button button9;
        private System.Windows.Forms.Button button10;
        private System.Windows.Forms.Button button11;
        private System.Windows.Forms.Button button12;
        private System.Windows.Forms.Button button13;
    }
}
