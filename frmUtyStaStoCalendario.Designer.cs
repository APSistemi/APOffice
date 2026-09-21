namespace APOffice
{
    partial class frmUtyStaStoCalendario
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
            this.dgv1 = new APOffice.APDataGridView();
            this.panel1 = new System.Windows.Forms.Panel();
            this.cmbYea = new System.Windows.Forms.ComboBox();
            this.label3 = new System.Windows.Forms.Label();
            this.btnSave = new System.Windows.Forms.Button();
            this.btnFill = new System.Windows.Forms.Button();
            this.panel2 = new System.Windows.Forms.Panel();
            this.label1 = new System.Windows.Forms.Label();
            this.dtpDti = new System.Windows.Forms.DateTimePicker();
            this.dtpDtf = new System.Windows.Forms.DateTimePicker();
            this.label2 = new System.Windows.Forms.Label();
            this.lblSta = new System.Windows.Forms.Label();
            this.btnSta = new System.Windows.Forms.Button();
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
            this.esciToolStripMenuItem});
            this.menuStrip1.Location = new System.Drawing.Point(0, 0);
            this.menuStrip1.Name = "menuStrip1";
            this.menuStrip1.Size = new System.Drawing.Size(685, 24);
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
            this.dgv1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgv1.Location = new System.Drawing.Point(35, 164);
            this.dgv1.Name = "dgv1";
            this.dgv1.Size = new System.Drawing.Size(617, 440);
            this.dgv1.TabIndex = 1;
            this.dgv1.CellDoubleClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dgv1_CellDoubleClick);
            // 
            // panel1
            // 
            this.panel1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.panel1.Controls.Add(this.cmbYea);
            this.panel1.Controls.Add(this.label3);
            this.panel1.Controls.Add(this.btnSave);
            this.panel1.Controls.Add(this.btnFill);
            this.panel1.Location = new System.Drawing.Point(0, 26);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(685, 63);
            this.panel1.TabIndex = 12;
            // 
            // cmbYea
            // 
            this.cmbYea.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.cmbYea.FormattingEnabled = true;
            this.cmbYea.Location = new System.Drawing.Point(276, 18);
            this.cmbYea.Name = "cmbYea";
            this.cmbYea.Size = new System.Drawing.Size(132, 28);
            this.cmbYea.TabIndex = 15;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(238, 30);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(32, 13);
            this.label3.TabIndex = 14;
            this.label3.Text = "Anno";
            // 
            // btnSave
            // 
            this.btnSave.Location = new System.Drawing.Point(492, 16);
            this.btnSave.Name = "btnSave";
            this.btnSave.Size = new System.Drawing.Size(142, 28);
            this.btnSave.TabIndex = 13;
            this.btnSave.Text = "Salva";
            this.btnSave.UseVisualStyleBackColor = true;
            this.btnSave.Click += new System.EventHandler(this.btnSave_Click);
            // 
            // btnFill
            // 
            this.btnFill.Location = new System.Drawing.Point(17, 16);
            this.btnFill.Name = "btnFill";
            this.btnFill.Size = new System.Drawing.Size(142, 28);
            this.btnFill.TabIndex = 12;
            this.btnFill.Text = "Inserisci date";
            this.btnFill.UseVisualStyleBackColor = true;
            this.btnFill.Click += new System.EventHandler(this.btnFill_Click);
            // 
            // panel2
            // 
            this.panel2.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.panel2.Controls.Add(this.btnSta);
            this.panel2.Controls.Add(this.lblSta);
            this.panel2.Controls.Add(this.label2);
            this.panel2.Controls.Add(this.dtpDtf);
            this.panel2.Controls.Add(this.dtpDti);
            this.panel2.Controls.Add(this.label1);
            this.panel2.Location = new System.Drawing.Point(0, 89);
            this.panel2.Name = "panel2";
            this.panel2.Size = new System.Drawing.Size(685, 58);
            this.panel2.TabIndex = 13;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(8, 10);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(88, 13);
            this.label1.TabIndex = 0;
            this.label1.Text = "Cambio stato dal ";
            // 
            // dtpDti
            // 
            this.dtpDti.CalendarFont = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.dtpDti.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.dtpDti.Location = new System.Drawing.Point(11, 26);
            this.dtpDti.Name = "dtpDti";
            this.dtpDti.Size = new System.Drawing.Size(232, 26);
            this.dtpDti.TabIndex = 1;
            // 
            // dtpDtf
            // 
            this.dtpDtf.CalendarFont = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.dtpDtf.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.dtpDtf.Location = new System.Drawing.Point(249, 26);
            this.dtpDtf.Name = "dtpDtf";
            this.dtpDtf.Size = new System.Drawing.Size(229, 26);
            this.dtpDtf.TabIndex = 2;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(251, 10);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(15, 13);
            this.label2.TabIndex = 3;
            this.label2.Text = "al";
            // 
            // lblSta
            // 
            this.lblSta.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.lblSta.Font = new System.Drawing.Font("Microsoft Sans Serif", 15.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblSta.Location = new System.Drawing.Point(497, 10);
            this.lblSta.Name = "lblSta";
            this.lblSta.Size = new System.Drawing.Size(55, 39);
            this.lblSta.TabIndex = 4;
            this.lblSta.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            this.lblSta.Click += new System.EventHandler(this.lblSta_Click);
            // 
            // btnSta
            // 
            this.btnSta.Location = new System.Drawing.Point(567, 10);
            this.btnSta.Name = "btnSta";
            this.btnSta.Size = new System.Drawing.Size(104, 39);
            this.btnSta.TabIndex = 14;
            this.btnSta.Text = "Cambia stato";
            this.btnSta.UseVisualStyleBackColor = true;
            this.btnSta.Click += new System.EventHandler(this.btnSta_Click);
            // 
            // frmUtyStaStoCalendario
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(685, 615);
            this.ControlBox = false;
            this.Controls.Add(this.panel2);
            this.Controls.Add(this.panel1);
            this.Controls.Add(this.dgv1);
            this.Controls.Add(this.menuStrip1);
            this.KeyPreview = true;
            this.MainMenuStrip = this.menuStrip1;
            this.Name = "frmUtyStaStoCalendario";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Calendario statistiche storico";
            this.Load += new System.EventHandler(this.frmUtyStaStoCalendario_Load);
            this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmUtyStaStoCalendario_KeyDown);
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
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.ComboBox cmbYea;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Button btnSave;
        private System.Windows.Forms.Button btnFill;
        private System.Windows.Forms.Panel panel2;
        private System.Windows.Forms.Label lblSta;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.DateTimePicker dtpDtf;
        private System.Windows.Forms.DateTimePicker dtpDti;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button btnSta;
    }
}
