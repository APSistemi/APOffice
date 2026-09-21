namespace APOffice
{
    partial class frmAnaArtImgWebSearch
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
            this.panelTop = new System.Windows.Forms.Panel();
            this.lblHeaderTitle = new System.Windows.Forms.Label();
            this.lblQuery = new System.Windows.Forms.Label();
            this.cmbQuery = new System.Windows.Forms.ComboBox();
            this.btnCerca = new System.Windows.Forms.Button();
            this.progressBarSearch = new System.Windows.Forms.ProgressBar();
            this.lblSearchStatus = new System.Windows.Forms.Label();
            this.splitContainer1 = new System.Windows.Forms.SplitContainer();
            this.flowGallery = new System.Windows.Forms.FlowLayoutPanel();
            this.panelPreview = new System.Windows.Forms.Panel();
            this.lblPreviewTitle = new System.Windows.Forms.Label();
            this.picBigPreview = new System.Windows.Forms.PictureBox();
            this.lblPreviewInfo = new System.Windows.Forms.Label();
            this.btnAbbina = new System.Windows.Forms.Button();
            this.btnApriBrowser = new System.Windows.Forms.Button();
            this.panelBottom = new System.Windows.Forms.Panel();
            this.lblStatusBottom = new System.Windows.Forms.Label();
            this.btnChiudi = new System.Windows.Forms.Button();
            this.panelTop.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.splitContainer1)).BeginInit();
            this.splitContainer1.Panel1.SuspendLayout();
            this.splitContainer1.Panel2.SuspendLayout();
            this.splitContainer1.SuspendLayout();
            this.panelPreview.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.picBigPreview)).BeginInit();
            this.panelBottom.SuspendLayout();
            this.SuspendLayout();
            // 
            // panelTop
            // 
            this.panelTop.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(15)))), ((int)(((byte)(23)))), ((int)(((byte)(42)))));
            this.panelTop.Controls.Add(this.lblHeaderTitle);
            this.panelTop.Controls.Add(this.lblQuery);
            this.panelTop.Controls.Add(this.cmbQuery);
            this.panelTop.Controls.Add(this.btnCerca);
            this.panelTop.Controls.Add(this.progressBarSearch);
            this.panelTop.Controls.Add(this.lblSearchStatus);
            this.panelTop.Dock = System.Windows.Forms.DockStyle.Top;
            this.panelTop.Location = new System.Drawing.Point(0, 0);
            this.panelTop.Name = "panelTop";
            this.panelTop.Size = new System.Drawing.Size(984, 90);
            this.panelTop.TabIndex = 0;
            // 
            // lblHeaderTitle
            // 
            this.lblHeaderTitle.AutoSize = true;
            this.lblHeaderTitle.Font = new System.Drawing.Font("Segoe UI", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblHeaderTitle.ForeColor = System.Drawing.Color.White;
            this.lblHeaderTitle.Location = new System.Drawing.Point(12, 9);
            this.lblHeaderTitle.Name = "lblHeaderTitle";
            this.lblHeaderTitle.Size = new System.Drawing.Size(280, 21);
            this.lblHeaderTitle.TabIndex = 0;
            this.lblHeaderTitle.Text = "🔍 Ricerca Foto Prodotto su Internet";
            // 
            // lblQuery
            // 
            this.lblQuery.AutoSize = true;
            this.lblQuery.Font = new System.Drawing.Font("Segoe UI", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblQuery.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(226)))), ((int)(((byte)(232)))), ((int)(((byte)(240)))));
            this.lblQuery.Location = new System.Drawing.Point(13, 38);
            this.lblQuery.Name = "lblQuery";
            this.lblQuery.Size = new System.Drawing.Size(147, 17);
            this.lblQuery.TabIndex = 1;
            this.lblQuery.Text = "Barcode / Descrizione:";
            // 
            // cmbQuery
            // 
            this.cmbQuery.Font = new System.Drawing.Font("Segoe UI", 10.5F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.cmbQuery.FormattingEnabled = true;
            this.cmbQuery.Location = new System.Drawing.Point(165, 34);
            this.cmbQuery.Name = "cmbQuery";
            this.cmbQuery.Size = new System.Drawing.Size(630, 27);
            this.cmbQuery.TabIndex = 2;
            this.cmbQuery.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbQuery_KeyDown);
            // 
            // btnCerca
            // 
            this.btnCerca.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(37)))), ((int)(((byte)(99)))), ((int)(((byte)(235)))));
            this.btnCerca.Cursor = System.Windows.Forms.Cursors.Hand;
            this.btnCerca.FlatAppearance.BorderSize = 0;
            this.btnCerca.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnCerca.Font = new System.Drawing.Font("Segoe UI", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnCerca.ForeColor = System.Drawing.Color.White;
            this.btnCerca.Location = new System.Drawing.Point(805, 32);
            this.btnCerca.Name = "btnCerca";
            this.btnCerca.Size = new System.Drawing.Size(165, 30);
            this.btnCerca.TabIndex = 3;
            this.btnCerca.Text = "🔍 Cerca Immagini";
            this.btnCerca.UseVisualStyleBackColor = false;
            this.btnCerca.Click += new System.EventHandler(this.btnCerca_Click);
            // 
            // progressBarSearch
            // 
            this.progressBarSearch.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.progressBarSearch.Location = new System.Drawing.Point(0, 86);
            this.progressBarSearch.Name = "progressBarSearch";
            this.progressBarSearch.Size = new System.Drawing.Size(984, 4);
            this.progressBarSearch.Style = System.Windows.Forms.ProgressBarStyle.Marquee;
            this.progressBarSearch.TabIndex = 4;
            this.progressBarSearch.Visible = false;
            // 
            // lblSearchStatus
            // 
            this.lblSearchStatus.AutoSize = true;
            this.lblSearchStatus.Font = new System.Drawing.Font("Segoe UI", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblSearchStatus.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(148)))), ((int)(((byte)(163)))), ((int)(((byte)(184)))));
            this.lblSearchStatus.Location = new System.Drawing.Point(165, 65);
            this.lblSearchStatus.Name = "lblSearchStatus";
            this.lblSearchStatus.Size = new System.Drawing.Size(351, 15);
            this.lblSearchStatus.TabIndex = 5;
            this.lblSearchStatus.Text = "Seleziona o inserisci un Barcode / Descrizione e premi Cerca Immagini";
            // 
            // splitContainer1
            // 
            this.splitContainer1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.splitContainer1.FixedPanel = System.Windows.Forms.FixedPanel.Panel2;
            this.splitContainer1.Location = new System.Drawing.Point(0, 90);
            this.splitContainer1.Name = "splitContainer1";
            // 
            // splitContainer1.Panel1
            // 
            this.splitContainer1.Panel1.Controls.Add(this.flowGallery);
            // 
            // splitContainer1.Panel2
            // 
            this.splitContainer1.Panel2.BackColor = System.Drawing.Color.White;
            this.splitContainer1.Panel2.Controls.Add(this.panelPreview);
            this.splitContainer1.Size = new System.Drawing.Size(984, 470);
            this.splitContainer1.SplitterDistance = 640;
            this.splitContainer1.TabIndex = 1;
            // 
            // flowGallery
            // 
            this.flowGallery.AutoScroll = true;
            this.flowGallery.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(241)))), ((int)(((byte)(245)))), ((int)(((byte)(249)))));
            this.flowGallery.Dock = System.Windows.Forms.DockStyle.Fill;
            this.flowGallery.Location = new System.Drawing.Point(0, 0);
            this.flowGallery.Name = "flowGallery";
            this.flowGallery.Padding = new System.Windows.Forms.Padding(10);
            this.flowGallery.Size = new System.Drawing.Size(640, 470);
            this.flowGallery.TabIndex = 0;
            // 
            // panelPreview
            // 
            this.panelPreview.BackColor = System.Drawing.Color.White;
            this.panelPreview.Controls.Add(this.lblPreviewTitle);
            this.panelPreview.Controls.Add(this.picBigPreview);
            this.panelPreview.Controls.Add(this.lblPreviewInfo);
            this.panelPreview.Controls.Add(this.btnAbbina);
            this.panelPreview.Controls.Add(this.btnApriBrowser);
            this.panelPreview.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panelPreview.Location = new System.Drawing.Point(0, 0);
            this.panelPreview.Name = "panelPreview";
            this.panelPreview.Padding = new System.Windows.Forms.Padding(15);
            this.panelPreview.Size = new System.Drawing.Size(340, 470);
            this.panelPreview.TabIndex = 0;
            // 
            // lblPreviewTitle
            // 
            this.lblPreviewTitle.AutoSize = true;
            this.lblPreviewTitle.Font = new System.Drawing.Font("Segoe UI", 11.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblPreviewTitle.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(15)))), ((int)(((byte)(23)))), ((int)(((byte)(42)))));
            this.lblPreviewTitle.Location = new System.Drawing.Point(12, 12);
            this.lblPreviewTitle.Name = "lblPreviewTitle";
            this.lblPreviewTitle.Size = new System.Drawing.Size(157, 20);
            this.lblPreviewTitle.TabIndex = 0;
            this.lblPreviewTitle.Text = "Anteprima Immagine";
            // 
            // picBigPreview
            // 
            this.picBigPreview.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(248)))), ((int)(((byte)(250)))), ((int)(((byte)(252)))));
            this.picBigPreview.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.picBigPreview.Location = new System.Drawing.Point(15, 40);
            this.picBigPreview.Name = "picBigPreview";
            this.picBigPreview.Size = new System.Drawing.Size(310, 240);
            this.picBigPreview.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom;
            this.picBigPreview.TabIndex = 1;
            this.picBigPreview.TabStop = false;
            // 
            // lblPreviewInfo
            // 
            this.lblPreviewInfo.Font = new System.Drawing.Font("Segoe UI", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblPreviewInfo.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(71)))), ((int)(((byte)(85)))), ((int)(((byte)(105)))));
            this.lblPreviewInfo.Location = new System.Drawing.Point(15, 290);
            this.lblPreviewInfo.Name = "lblPreviewInfo";
            this.lblPreviewInfo.Size = new System.Drawing.Size(310, 80);
            this.lblPreviewInfo.TabIndex = 2;
            this.lblPreviewInfo.Text = "Nessuna immagine selezionata.\nClicca su una foto nella galleria a sinistra per vi" +
    "sualizzarla in dettaglio.";
            // 
            // btnAbbina
            // 
            this.btnAbbina.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(22)))), ((int)(((byte)(163)))), ((int)(((byte)(74)))));
            this.btnAbbina.Cursor = System.Windows.Forms.Cursors.Hand;
            this.btnAbbina.Enabled = false;
            this.btnAbbina.FlatAppearance.BorderSize = 0;
            this.btnAbbina.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnAbbina.Font = new System.Drawing.Font("Segoe UI", 10.5F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnAbbina.ForeColor = System.Drawing.Color.White;
            this.btnAbbina.Location = new System.Drawing.Point(15, 380);
            this.btnAbbina.Name = "btnAbbina";
            this.btnAbbina.Size = new System.Drawing.Size(310, 40);
            this.btnAbbina.TabIndex = 3;
            this.btnAbbina.Text = "✅ Abbina a Questo Articolo";
            this.btnAbbina.UseVisualStyleBackColor = false;
            this.btnAbbina.Click += new System.EventHandler(this.btnAbbina_Click);
            // 
            // btnApriBrowser
            // 
            this.btnApriBrowser.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(241)))), ((int)(((byte)(245)))), ((int)(((byte)(249)))));
            this.btnApriBrowser.Cursor = System.Windows.Forms.Cursors.Hand;
            this.btnApriBrowser.Enabled = false;
            this.btnApriBrowser.FlatAppearance.BorderColor = System.Drawing.Color.FromArgb(((int)(((byte)(203)))), ((int)(((byte)(213)))), ((int)(((byte)(225)))));
            this.btnApriBrowser.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnApriBrowser.Font = new System.Drawing.Font("Segoe UI", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnApriBrowser.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(51)))), ((int)(((byte)(65)))), ((int)(((byte)(85)))));
            this.btnApriBrowser.Location = new System.Drawing.Point(15, 427);
            this.btnApriBrowser.Name = "btnApriBrowser";
            this.btnApriBrowser.Size = new System.Drawing.Size(310, 28);
            this.btnApriBrowser.TabIndex = 4;
            this.btnApriBrowser.Text = "🌐 Visualizza a Dimensione Massima";
            this.btnApriBrowser.UseVisualStyleBackColor = false;
            this.btnApriBrowser.Click += new System.EventHandler(this.btnApriBrowser_Click);
            // 
            // panelBottom
            // 
            this.panelBottom.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(248)))), ((int)(((byte)(250)))), ((int)(((byte)(252)))));
            this.panelBottom.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.panelBottom.Controls.Add(this.lblStatusBottom);
            this.panelBottom.Controls.Add(this.btnChiudi);
            this.panelBottom.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.panelBottom.Location = new System.Drawing.Point(0, 560);
            this.panelBottom.Name = "panelBottom";
            this.panelBottom.Size = new System.Drawing.Size(984, 44);
            this.panelBottom.TabIndex = 2;
            // 
            // lblStatusBottom
            // 
            this.lblStatusBottom.AutoSize = true;
            this.lblStatusBottom.Font = new System.Drawing.Font("Segoe UI", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblStatusBottom.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(51)))), ((int)(((byte)(65)))), ((int)(((byte)(85)))));
            this.lblStatusBottom.Location = new System.Drawing.Point(12, 12);
            this.lblStatusBottom.Name = "lblStatusBottom";
            this.lblStatusBottom.Size = new System.Drawing.Size(183, 17);
            this.lblStatusBottom.TabIndex = 0;
            this.lblStatusBottom.Text = "Nessuna immagine caricata.";
            // 
            // btnChiudi
            // 
            this.btnChiudi.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btnChiudi.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(226)))), ((int)(((byte)(232)))), ((int)(((byte)(240)))));
            this.btnChiudi.Cursor = System.Windows.Forms.Cursors.Hand;
            this.btnChiudi.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.btnChiudi.FlatAppearance.BorderSize = 0;
            this.btnChiudi.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnChiudi.Font = new System.Drawing.Font("Segoe UI", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnChiudi.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(30)))), ((int)(((byte)(41)))), ((int)(((byte)(59)))));
            this.btnChiudi.Location = new System.Drawing.Point(867, 7);
            this.btnChiudi.Name = "btnChiudi";
            this.btnChiudi.Size = new System.Drawing.Size(104, 28);
            this.btnChiudi.TabIndex = 1;
            this.btnChiudi.Text = "Annulla (ESC)";
            this.btnChiudi.UseVisualStyleBackColor = false;
            this.btnChiudi.Click += new System.EventHandler(this.btnChiudi_Click);
            // 
            // frmAnaArtImgWebSearch
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(248)))), ((int)(((byte)(250)))), ((int)(((byte)(252)))));
            this.CancelButton = this.btnChiudi;
            this.ClientSize = new System.Drawing.Size(984, 604);
            this.Controls.Add(this.splitContainer1);
            this.Controls.Add(this.panelBottom);
            this.Controls.Add(this.panelTop);
            this.Font = new System.Drawing.Font("Segoe UI", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.MinimumSize = new System.Drawing.Size(800, 500);
            this.Name = "frmAnaArtImgWebSearch";
            this.ShowIcon = false;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "Ricerca e Abbinamento Foto Prodotto da Web";
            this.Load += new System.EventHandler(this.frmAnaArtImgWebSearch_Load);
            this.panelTop.ResumeLayout(false);
            this.panelTop.PerformLayout();
            this.splitContainer1.Panel1.ResumeLayout(false);
            this.splitContainer1.Panel2.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.splitContainer1)).EndInit();
            this.splitContainer1.ResumeLayout(false);
            this.panelPreview.ResumeLayout(false);
            this.panelPreview.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.picBigPreview)).EndInit();
            this.panelBottom.ResumeLayout(false);
            this.panelBottom.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Panel panelTop;
        private System.Windows.Forms.Label lblHeaderTitle;
        private System.Windows.Forms.Label lblQuery;
        private System.Windows.Forms.ComboBox cmbQuery;
        private System.Windows.Forms.Button btnCerca;
        private System.Windows.Forms.ProgressBar progressBarSearch;
        private System.Windows.Forms.Label lblSearchStatus;
        private System.Windows.Forms.SplitContainer splitContainer1;
        private System.Windows.Forms.FlowLayoutPanel flowGallery;
        private System.Windows.Forms.Panel panelPreview;
        private System.Windows.Forms.Label lblPreviewTitle;
        private System.Windows.Forms.PictureBox picBigPreview;
        private System.Windows.Forms.Label lblPreviewInfo;
        private System.Windows.Forms.Button btnAbbina;
        private System.Windows.Forms.Button btnApriBrowser;
        private System.Windows.Forms.Panel panelBottom;
        private System.Windows.Forms.Label lblStatusBottom;
        private System.Windows.Forms.Button btnChiudi;
    }
}
