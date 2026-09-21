namespace APOffice
{
    partial class frmGesChiusure
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
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle2 = new System.Windows.Forms.DataGridViewCellStyle();
            this.menuStrip1 = new System.Windows.Forms.MenuStrip();
            this.esciToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.dgv1 = new APOffice.APDataGridView();
            this.btnChiudi = new System.Windows.Forms.Button();
            this.progressBar1 = new System.Windows.Forms.ProgressBar();
            this.lblMsg = new System.Windows.Forms.Label();
            this.dgv2 = new APOffice.APDataGridView();
            this.btnPrnRep = new System.Windows.Forms.Button();
            this.dtpVenRep = new System.Windows.Forms.DateTimePicker();
            this.btnVen = new System.Windows.Forms.Button();
            this.lblTot = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.chkVenClo = new System.Windows.Forms.CheckBox();
            this.chkPer100 = new System.Windows.Forms.CheckBox();
            this.menuStrip1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgv1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.dgv2)).BeginInit();
            this.SuspendLayout();
            // 
            // menuStrip1
            // 
            this.menuStrip1.BackColor = System.Drawing.SystemColors.ControlLight;
            this.menuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.esciToolStripMenuItem});
            this.menuStrip1.Location = new System.Drawing.Point(0, 0);
            this.menuStrip1.Name = "menuStrip1";
            this.menuStrip1.Size = new System.Drawing.Size(664, 24);
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
            // dgv1
            // 
            dataGridViewCellStyle1.BackColor = System.Drawing.Color.Lavender;
            this.dgv1.AlternatingRowsDefaultCellStyle = dataGridViewCellStyle1;
            this.dgv1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgv1.Location = new System.Drawing.Point(12, 62);
            this.dgv1.Name = "dgv1";
            this.dgv1.Size = new System.Drawing.Size(466, 91);
            this.dgv1.TabIndex = 1;
            // 
            // btnChiudi
            // 
            this.btnChiudi.Location = new System.Drawing.Point(484, 27);
            this.btnChiudi.Name = "btnChiudi";
            this.btnChiudi.Size = new System.Drawing.Size(164, 29);
            this.btnChiudi.TabIndex = 2;
            this.btnChiudi.Text = "Avvio chiusura fine giorno";
            this.btnChiudi.UseVisualStyleBackColor = true;
            this.btnChiudi.Click += new System.EventHandler(this.btnChiudi_Click);
            // 
            // progressBar1
            // 
            this.progressBar1.Location = new System.Drawing.Point(3, 478);
            this.progressBar1.Name = "progressBar1";
            this.progressBar1.Size = new System.Drawing.Size(657, 12);
            this.progressBar1.TabIndex = 3;
            // 
            // lblMsg
            // 
            this.lblMsg.AutoSize = true;
            this.lblMsg.Location = new System.Drawing.Point(13, 457);
            this.lblMsg.Name = "lblMsg";
            this.lblMsg.Size = new System.Drawing.Size(16, 13);
            this.lblMsg.TabIndex = 4;
            this.lblMsg.Text = "...";
            // 
            // dgv2
            // 
            dataGridViewCellStyle2.BackColor = System.Drawing.Color.Lavender;
            this.dgv2.AlternatingRowsDefaultCellStyle = dataGridViewCellStyle2;
            this.dgv2.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgv2.Location = new System.Drawing.Point(12, 159);
            this.dgv2.Name = "dgv2";
            this.dgv2.Size = new System.Drawing.Size(636, 286);
            this.dgv2.TabIndex = 5;
            // 
            // btnPrnRep
            // 
            this.btnPrnRep.Location = new System.Drawing.Point(484, 62);
            this.btnPrnRep.Name = "btnPrnRep";
            this.btnPrnRep.Size = new System.Drawing.Size(164, 65);
            this.btnPrnRep.TabIndex = 6;
            this.btnPrnRep.Text = "Stampa";
            this.btnPrnRep.UseVisualStyleBackColor = true;
            this.btnPrnRep.Click += new System.EventHandler(this.btnPrnRep_Click);
            // 
            // dtpVenRep
            // 
            this.dtpVenRep.Location = new System.Drawing.Point(251, 32);
            this.dtpVenRep.Name = "dtpVenRep";
            this.dtpVenRep.Size = new System.Drawing.Size(186, 20);
            this.dtpVenRep.TabIndex = 7;
            // 
            // btnVen
            // 
            this.btnVen.Location = new System.Drawing.Point(12, 30);
            this.btnVen.Name = "btnVen";
            this.btnVen.Size = new System.Drawing.Size(233, 23);
            this.btnVen.TabIndex = 0;
            this.btnVen.Text = "Estrazione venduto per reparto da stampare al";
            this.btnVen.Click += new System.EventHandler(this.btnVen_Click);
            // 
            // lblTot
            // 
            this.lblTot.AutoSize = true;
            this.lblTot.Location = new System.Drawing.Point(570, 461);
            this.lblTot.Name = "lblTot";
            this.lblTot.Size = new System.Drawing.Size(13, 13);
            this.lblTot.TabIndex = 8;
            this.lblTot.Text = "0";
            this.lblTot.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(514, 461);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(37, 13);
            this.label1.TabIndex = 9;
            this.label1.Text = "Totale";
            // 
            // chkVenClo
            // 
            this.chkVenClo.AutoSize = true;
            this.chkVenClo.Checked = true;
            this.chkVenClo.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chkVenClo.Location = new System.Drawing.Point(484, 136);
            this.chkVenClo.Name = "chkVenClo";
            this.chkVenClo.Size = new System.Drawing.Size(119, 17);
            this.chkVenClo.TabIndex = 10;
            this.chkVenClo.Text = "Con chiusura cassa";
            this.chkVenClo.UseVisualStyleBackColor = true;
            // 
            // chkPer100
            // 
            this.chkPer100.AutoSize = true;
            this.chkPer100.Checked = true;
            this.chkPer100.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chkPer100.Location = new System.Drawing.Point(599, 136);
            this.chkPer100.Name = "chkPer100";
            this.chkPer100.Size = new System.Drawing.Size(52, 17);
            this.chkPer100.TabIndex = 11;
            this.chkPer100.Text = "x 100";
            this.chkPer100.UseVisualStyleBackColor = true;
            // 
            // frmGesChiusure
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(664, 493);
            this.Controls.Add(this.chkPer100);
            this.Controls.Add(this.chkVenClo);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.lblTot);
            this.Controls.Add(this.btnVen);
            this.Controls.Add(this.dtpVenRep);
            this.Controls.Add(this.btnPrnRep);
            this.Controls.Add(this.dgv2);
            this.Controls.Add(this.lblMsg);
            this.Controls.Add(this.progressBar1);
            this.Controls.Add(this.btnChiudi);
            this.Controls.Add(this.dgv1);
            this.Controls.Add(this.menuStrip1);
            this.KeyPreview = true;
            this.MainMenuStrip = this.menuStrip1;
            this.Name = "frmGesChiusure";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Chiusure casse";
            this.Load += new System.EventHandler(this.frmGesChiusure_Load);
            this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmGesChiusure_KeyDown);
            this.menuStrip1.ResumeLayout(false);
            this.menuStrip1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgv1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.dgv2)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.MenuStrip menuStrip1;
        private System.Windows.Forms.ToolStripMenuItem esciToolStripMenuItem;
        private APOffice.APDataGridView dgv1;
        private System.Windows.Forms.Button btnChiudi;
        private System.Windows.Forms.ProgressBar progressBar1;
        private System.Windows.Forms.Label lblMsg;
        private APOffice.APDataGridView dgv2;
        private System.Windows.Forms.Button btnPrnRep;
        private System.Windows.Forms.DateTimePicker dtpVenRep;
        private System.Windows.Forms.Button btnVen;
        private System.Windows.Forms.Label lblTot;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.CheckBox chkVenClo;
        private System.Windows.Forms.CheckBox chkPer100;
    }
}
