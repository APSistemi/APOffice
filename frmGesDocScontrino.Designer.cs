namespace APOffice
{
    partial class frmGesDocScontrino
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
            this.dtpDay = new System.Windows.Forms.DateTimePicker();
            this.menuStrip1 = new System.Windows.Forms.MenuStrip();
            this.esciToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.label1 = new System.Windows.Forms.Label();
            this.nudPos = new System.Windows.Forms.NumericUpDown();
            this.label2 = new System.Windows.Forms.Label();
            this.nudNum = new System.Windows.Forms.NumericUpDown();
            this.btnOk = new System.Windows.Forms.Button();
            this.dgv1 = new APOffice.APDataGridView();
            this.rdbCauSco = new System.Windows.Forms.RadioButton();
            this.rdbCauPro = new System.Windows.Forms.RadioButton();
            this.menuStrip1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.nudPos)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.nudNum)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.dgv1)).BeginInit();
            this.SuspendLayout();
            // 
            // dtpDay
            // 
            this.dtpDay.CalendarFont = new System.Drawing.Font("Microsoft Sans Serif", 20.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.dtpDay.Font = new System.Drawing.Font("Microsoft Sans Serif", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.dtpDay.Location = new System.Drawing.Point(6, 43);
            this.dtpDay.Name = "dtpDay";
            this.dtpDay.Size = new System.Drawing.Size(333, 29);
            this.dtpDay.TabIndex = 0;
            // 
            // menuStrip1
            // 
            this.menuStrip1.BackColor = System.Drawing.SystemColors.ControlLight;
            this.menuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.esciToolStripMenuItem});
            this.menuStrip1.Location = new System.Drawing.Point(0, 0);
            this.menuStrip1.Name = "menuStrip1";
            this.menuStrip1.Size = new System.Drawing.Size(344, 24);
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
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(12, 139);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(78, 20);
            this.label1.TabIndex = 2;
            this.label1.Text = "N. cassa";
            // 
            // nudPos
            // 
            this.nudPos.Font = new System.Drawing.Font("Microsoft Sans Serif", 21.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.nudPos.Location = new System.Drawing.Point(255, 130);
            this.nudPos.Name = "nudPos";
            this.nudPos.Size = new System.Drawing.Size(84, 40);
            this.nudPos.TabIndex = 3;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label2.Location = new System.Drawing.Point(12, 196);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(105, 20);
            this.label2.TabIndex = 4;
            this.label2.Text = "N. scontrino";
            // 
            // nudNum
            // 
            this.nudNum.Font = new System.Drawing.Font("Microsoft Sans Serif", 21.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.nudNum.Location = new System.Drawing.Point(255, 176);
            this.nudNum.Maximum = new decimal(new int[] {
            999,
            0,
            0,
            0});
            this.nudNum.Name = "nudNum";
            this.nudNum.Size = new System.Drawing.Size(84, 40);
            this.nudNum.TabIndex = 5;
            // 
            // btnOk
            // 
            this.btnOk.Location = new System.Drawing.Point(5, 222);
            this.btnOk.Name = "btnOk";
            this.btnOk.Size = new System.Drawing.Size(336, 27);
            this.btnOk.TabIndex = 6;
            this.btnOk.Text = "Estrai";
            this.btnOk.UseVisualStyleBackColor = true;
            this.btnOk.Click += new System.EventHandler(this.btnOk_Click);
            // 
            // dgv1
            // 
            dataGridViewCellStyle1.BackColor = System.Drawing.Color.Lavender;
            this.dgv1.AlternatingRowsDefaultCellStyle = dataGridViewCellStyle1;
            this.dgv1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgv1.Location = new System.Drawing.Point(7, 255);
            this.dgv1.Name = "dgv1";
            this.dgv1.RowHeadersWidth = 20;
            this.dgv1.Size = new System.Drawing.Size(332, 103);
            this.dgv1.TabIndex = 7;
            this.dgv1.DoubleClick += new System.EventHandler(this.dgv1_DoubleClick);
            this.dgv1.KeyDown += new System.Windows.Forms.KeyEventHandler(this.dgv1_KeyDown);
            // 
            // rdbCauSco
            // 
            this.rdbCauSco.AutoSize = true;
            this.rdbCauSco.Checked = true;
            this.rdbCauSco.Font = new System.Drawing.Font("Microsoft Sans Serif", 15.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.rdbCauSco.Location = new System.Drawing.Point(199, 88);
            this.rdbCauSco.Name = "rdbCauSco";
            this.rdbCauSco.Size = new System.Drawing.Size(130, 29);
            this.rdbCauSco.TabIndex = 8;
            this.rdbCauSco.TabStop = true;
            this.rdbCauSco.Text = "Scontrino";
            this.rdbCauSco.UseVisualStyleBackColor = true;
            // 
            // rdbCauPro
            // 
            this.rdbCauPro.AutoSize = true;
            this.rdbCauPro.CheckAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.rdbCauPro.Font = new System.Drawing.Font("Microsoft Sans Serif", 15.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.rdbCauPro.Location = new System.Drawing.Point(28, 88);
            this.rdbCauPro.Name = "rdbCauPro";
            this.rdbCauPro.Size = new System.Drawing.Size(125, 29);
            this.rdbCauPro.TabIndex = 9;
            this.rdbCauPro.Text = "Proforma";
            this.rdbCauPro.UseVisualStyleBackColor = true;
            // 
            // frmGesDocScontrino
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(344, 365);
            this.ControlBox = false;
            this.Controls.Add(this.rdbCauPro);
            this.Controls.Add(this.rdbCauSco);
            this.Controls.Add(this.dgv1);
            this.Controls.Add(this.btnOk);
            this.Controls.Add(this.nudNum);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.nudPos);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.dtpDay);
            this.Controls.Add(this.menuStrip1);
            this.KeyPreview = true;
            this.MainMenuStrip = this.menuStrip1;
            this.Name = "frmGesDocScontrino";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Ricerca Scontrino";
            this.Load += new System.EventHandler(this.frmGesDocScontrino_Load);
            this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmGesDocScontrino_KeyDown);
            this.menuStrip1.ResumeLayout(false);
            this.menuStrip1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.nudPos)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.nudNum)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.dgv1)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.DateTimePicker dtpDay;
        private System.Windows.Forms.MenuStrip menuStrip1;
        private System.Windows.Forms.ToolStripMenuItem esciToolStripMenuItem;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.NumericUpDown nudPos;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.NumericUpDown nudNum;
        private System.Windows.Forms.Button btnOk;
        private APOffice.APDataGridView dgv1;
        private System.Windows.Forms.RadioButton rdbCauSco;
        private System.Windows.Forms.RadioButton rdbCauPro;
    }
}
