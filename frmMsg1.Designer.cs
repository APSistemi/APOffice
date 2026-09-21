namespace APOffice
{
    partial class frmMsg1
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
            this.btnOk = new System.Windows.Forms.Button();
            this.btnNo = new System.Windows.Forms.Button();
            this.btnSi = new System.Windows.Forms.Button();
            this.lblMsg1 = new System.Windows.Forms.Label();
            this.chkMem = new System.Windows.Forms.CheckBox();
            this.SuspendLayout();
            // 
            // btnOk
            // 
            this.btnOk.BackColor = System.Drawing.Color.PowderBlue;
            this.btnOk.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnOk.Location = new System.Drawing.Point(202, 276);
            this.btnOk.Name = "btnOk";
            this.btnOk.Size = new System.Drawing.Size(116, 97);
            this.btnOk.TabIndex = 7;
            this.btnOk.Text = "OK";
            this.btnOk.UseVisualStyleBackColor = false;
            this.btnOk.Click += new System.EventHandler(this.btn_Click);
            // 
            // btnNo
            // 
            this.btnNo.BackColor = System.Drawing.Color.IndianRed;
            this.btnNo.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnNo.Location = new System.Drawing.Point(386, 276);
            this.btnNo.Name = "btnNo";
            this.btnNo.Size = new System.Drawing.Size(116, 97);
            this.btnNo.TabIndex = 6;
            this.btnNo.Text = "NO";
            this.btnNo.UseVisualStyleBackColor = false;
            this.btnNo.Click += new System.EventHandler(this.btn_Click);
            // 
            // btnSi
            // 
            this.btnSi.BackColor = System.Drawing.Color.PaleGreen;
            this.btnSi.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnSi.Location = new System.Drawing.Point(12, 276);
            this.btnSi.Name = "btnSi";
            this.btnSi.Size = new System.Drawing.Size(116, 97);
            this.btnSi.TabIndex = 5;
            this.btnSi.Text = "SI";
            this.btnSi.UseVisualStyleBackColor = false;
            this.btnSi.Click += new System.EventHandler(this.btn_Click);
            // 
            // lblMsg1
            // 
            this.lblMsg1.BackColor = System.Drawing.Color.White;
            this.lblMsg1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.lblMsg1.Font = new System.Drawing.Font("Arial Narrow", 14.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblMsg1.Location = new System.Drawing.Point(12, 9);
            this.lblMsg1.Name = "lblMsg1";
            this.lblMsg1.Size = new System.Drawing.Size(490, 244);
            this.lblMsg1.TabIndex = 4;
            this.lblMsg1.Text = "label1";
            this.lblMsg1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // chkMem
            // 
            this.chkMem.AutoSize = true;
            this.chkMem.Location = new System.Drawing.Point(12, 379);
            this.chkMem.Name = "chkMem";
            this.chkMem.Size = new System.Drawing.Size(202, 17);
            this.chkMem.TabIndex = 8;
            this.chkMem.Text = "Memorizza la scelta per tutti gli articoli";
            this.chkMem.UseVisualStyleBackColor = true;
            // 
            // frmMsg1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.AliceBlue;
            this.ClientSize = new System.Drawing.Size(518, 398);
            this.ControlBox = false;
            this.Controls.Add(this.chkMem);
            this.Controls.Add(this.btnOk);
            this.Controls.Add(this.btnNo);
            this.Controls.Add(this.btnSi);
            this.Controls.Add(this.lblMsg1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
            this.Name = "frmMsg1";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "frmMsg1";
            this.Load += new System.EventHandler(this.frmMsg1_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btnOk;
        private System.Windows.Forms.Button btnNo;
        private System.Windows.Forms.Button btnSi;
        private System.Windows.Forms.Label lblMsg1;
        private System.Windows.Forms.CheckBox chkMem;
    }
}