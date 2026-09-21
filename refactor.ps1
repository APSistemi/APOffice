$content = Get-Content "frmAnaArticolo.Designer.cs" -Raw

$decls = "        private System.Windows.Forms.TabControl tabMain;`r`n        private System.Windows.Forms.TabPage tabGenerale;`r`n        private System.Windows.Forms.TabPage tabBilancia;`r`n        private System.Windows.Forms.TabPage tabListini;"
$content = $content -replace 'private System\.Windows\.Forms\.Panel pnlArt;', ("private System.Windows.Forms.Panel pnlArt;`r`n" + $decls)

$inst = "            this.tabMain = new System.Windows.Forms.TabControl();`r`n            this.tabGenerale = new System.Windows.Forms.TabPage();`r`n            this.tabBilancia = new System.Windows.Forms.TabPage();`r`n            this.tabListini = new System.Windows.Forms.TabPage();"
$content = $content -replace 'this\.pnlArt = new System\.Windows\.Forms\.Panel\(\);', ("this.pnlArt = new System.Windows.Forms.Panel();`r`n" + $inst)

$susp = "            this.tabMain.SuspendLayout();`r`n            this.tabGenerale.SuspendLayout();`r`n            this.tabBilancia.SuspendLayout();`r`n            this.tabListini.SuspendLayout();"
$content = $content -replace 'this\.pnlArt\.SuspendLayout\(\);', ("this.pnlArt.SuspendLayout();`r`n" + $susp)

$content = $content -replace 'this\.pnlArt\.Controls\.Add\(this\.panel1\);', "this.tabGenerale.Controls.Add(this.panel1);"
$content = $content -replace 'this\.pnlArt\.Controls\.Add\(this\.groupBox2\);', "this.tabGenerale.Controls.Add(this.groupBox2);"

$content = $content -replace 'this\.pnlArt\.Controls\.Add\(this\.groupBox1\);', "this.tabBilancia.Controls.Add(this.groupBox1);"
$content = $content -replace 'this\.pnlArt\.Controls\.Add\(this\.groupBox3\);', "this.tabBilancia.Controls.Add(this.groupBox3);"

$content = $content -replace 'this\.pnlArt\.Controls\.Add\(this\.label3\);', "this.tabListini.Controls.Add(this.label3);"
$content = $content -replace 'this\.pnlArt\.Controls\.Add\(this\.txtEan\);', "this.tabListini.Controls.Add(this.txtEan);"
$content = $content -replace 'this\.pnlArt\.Controls\.Add\(this\.label38\);', "this.tabListini.Controls.Add(this.label38);"
$content = $content -replace 'this\.pnlArt\.Controls\.Add\(this\.btnEanAutomatico\);', "this.tabListini.Controls.Add(this.btnEanAutomatico);"
$content = $content -replace 'this\.pnlArt\.Controls\.Add\(this\.dgv1\);', "this.tabListini.Controls.Add(this.dgv1);"

$content = $content -replace 'this\.pnlArt\.Controls\.Add\(this\.label1\);', "this.tabListini.Controls.Add(this.label1);"
$content = $content -replace 'this\.pnlArt\.Controls\.Add\(this\.dgv2\);', "this.tabListini.Controls.Add(this.dgv2);"
$content = $content -replace 'this\.pnlArt\.Controls\.Add\(this\.btnForCos\);', "this.tabListini.Controls.Add(this.btnForCos);"

$content = $content -replace 'this\.pnlArt\.Controls\.Add\(this\.label2\);', "this.tabListini.Controls.Add(this.label2);"
$content = $content -replace 'this\.pnlArt\.Controls\.Add\(this\.dgv3\);', "this.tabListini.Controls.Add(this.dgv3);"
$content = $content -replace 'this\.pnlArt\.Controls\.Add\(this\.label39\);', "this.tabListini.Controls.Add(this.label39);"
$content = $content -replace 'this\.pnlArt\.Controls\.Add\(this\.cmbLisVen\);', "this.tabListini.Controls.Add(this.cmbLisVen);"

$tab_config = @"
            // 
            // tabMain
            // 
            this.tabMain.Controls.Add(this.tabGenerale);
            this.tabMain.Controls.Add(this.tabBilancia);
            this.tabMain.Controls.Add(this.tabListini);
            this.tabMain.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tabMain.Font = new System.Drawing.Font("Segoe UI", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.tabMain.Location = new System.Drawing.Point(0, 0);
            this.tabMain.Name = "tabMain";
            this.tabMain.SelectedIndex = 0;
            this.tabMain.Size = new System.Drawing.Size(897, 545);
            this.tabMain.TabIndex = 2;
            // 
            // tabGenerale
            // 
            this.tabGenerale.BackColor = System.Drawing.Color.White;
            this.tabGenerale.Location = new System.Drawing.Point(4, 26);
            this.tabGenerale.Name = "tabGenerale";
            this.tabGenerale.Padding = new System.Windows.Forms.Padding(3);
            this.tabGenerale.Size = new System.Drawing.Size(889, 515);
            this.tabGenerale.TabIndex = 0;
            this.tabGenerale.Text = "Informazioni Generali";
            // 
            // tabBilancia
            // 
            this.tabBilancia.BackColor = System.Drawing.Color.White;
            this.tabBilancia.Location = new System.Drawing.Point(4, 26);
            this.tabBilancia.Name = "tabBilancia";
            this.tabBilancia.Padding = new System.Windows.Forms.Padding(3);
            this.tabBilancia.Size = new System.Drawing.Size(889, 515);
            this.tabBilancia.TabIndex = 1;
            this.tabBilancia.Text = "Dati Bilancia ed Economici";
            // 
            // tabListini
            // 
            this.tabListini.BackColor = System.Drawing.Color.White;
            this.tabListini.Location = new System.Drawing.Point(4, 26);
            this.tabListini.Name = "tabListini";
            this.tabListini.Padding = new System.Windows.Forms.Padding(3);
            this.tabListini.Size = new System.Drawing.Size(889, 515);
            this.tabListini.TabIndex = 2;
            this.tabListini.Text = "Barcode e Listini";
"@
$content = $content -replace 'this\.pnlArt\.TabIndex = 1;', ("this.pnlArt.TabIndex = 1;`r`n            this.pnlArt.Controls.Add(this.tabMain);`r`n" + $tab_config)

$content = $content -replace 'this\.panel1\.Location = new System\.Drawing\.Point\([^)]+\);', 'this.panel1.Location = new System.Drawing.Point(5, 5);'
$content = $content -replace 'this\.groupBox2\.Location = new System\.Drawing\.Point\([^)]+\);', 'this.groupBox2.Location = new System.Drawing.Point(5, 155);'

$content = $content -replace 'this\.groupBox1\.Location = new System\.Drawing\.Point\([^)]+\);', 'this.groupBox1.Location = new System.Drawing.Point(10, 10);'
$content = $content -replace 'this\.groupBox1\.Size = new System\.Drawing\.Size\([^)]+\);', 'this.groupBox1.Size = new System.Drawing.Size(275, 480);'

$content = $content -replace 'this\.groupBox3\.Location = new System\.Drawing\.Point\([^)]+\);', 'this.groupBox3.Location = new System.Drawing.Point(300, 10);'
$content = $content -replace 'this\.groupBox3\.Size = new System\.Drawing\.Size\([^)]+\);', 'this.groupBox3.Size = new System.Drawing.Size(330, 480);'

$content = $content -replace 'this\.label3\.Location = new System\.Drawing\.Point\([^)]+\);', 'this.label3.Location = new System.Drawing.Point(10, 15);'
$content = $content -replace 'this\.txtEan\.Location = new System\.Drawing\.Point\([^)]+\);', 'this.txtEan.Location = new System.Drawing.Point(70, 12);'
$content = $content -replace 'this\.label38\.Location = new System\.Drawing\.Point\([^)]+\);', 'this.label38.Location = new System.Drawing.Point(205, 15);'
$content = $content -replace 'this\.btnEanAutomatico\.Location = new System\.Drawing\.Point\([^)]+\);', 'this.btnEanAutomatico.Location = new System.Drawing.Point(250, 10);'
$content = $content -replace 'this\.dgv1\.Location = new System\.Drawing\.Point\([^)]+\);', 'this.dgv1.Location = new System.Drawing.Point(10, 40);'
$content = $content -replace 'this\.dgv1\.Size = new System\.Drawing\.Size\([^)]+\);', 'this.dgv1.Size = new System.Drawing.Size(270, 460);'

$content = $content -replace 'this\.label1\.Location = new System\.Drawing\.Point\([^)]+\);', 'this.label1.Location = new System.Drawing.Point(295, 15);'
$content = $content -replace 'this\.btnForCos\.Location = new System\.Drawing\.Point\([^)]+\);', 'this.btnForCos.Location = new System.Drawing.Point(450, 10);'
$content = $content -replace 'this\.dgv2\.Location = new System\.Drawing\.Point\([^)]+\);', 'this.dgv2.Location = new System.Drawing.Point(295, 40);'
$content = $content -replace 'this\.dgv2\.Size = new System\.Drawing\.Size\([^)]+\);', 'this.dgv2.Size = new System.Drawing.Size(280, 460);'

$content = $content -replace 'this\.label2\.Location = new System\.Drawing\.Point\([^)]+\);', 'this.label2.Location = new System.Drawing.Point(590, 15);'
$content = $content -replace 'this\.label39\.Location = new System\.Drawing\.Point\([^)]+\);', 'this.label39.Location = new System.Drawing.Point(590, 15);'
$content = $content -replace 'this\.cmbLisVen\.Location = new System\.Drawing\.Point\([^)]+\);', 'this.cmbLisVen.Location = new System.Drawing.Point(700, 10);'
$content = $content -replace 'this\.dgv3\.Location = new System\.Drawing\.Point\([^)]+\);', 'this.dgv3.Location = new System.Drawing.Point(590, 40);'
$content = $content -replace 'this\.dgv3\.Size = new System\.Drawing\.Size\([^)]+\);', 'this.dgv3.Size = new System.Drawing.Size(280, 460);'

$content = $content -replace 'System\.Drawing\.Color\.FromArgb\(\(\(int\)\(\(\(byte\)\(224\)\)\)\), \(\(\(int\)\(\(\(byte\)\(224\)\)\)\), \(\(\(int\)\(\(\(byte\)\(224\)\)\)\)\)', 'System.Drawing.Color.WhiteSmoke'
$content = $content -replace 'this\.panel1\.BackColor = System\.Drawing\.Color\.WhiteSmoke;', 'this.panel1.BackColor = System.Drawing.Color.White;'

$content = $content -replace 'this\.Text = " Anagrafica articolo";', "this.Text = `" Anagrafica articolo`";`r`n            this.Font = new System.Drawing.Font(`"Segoe UI`", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));"

$resume = "            this.tabMain.ResumeLayout(false);`r`n            this.tabGenerale.ResumeLayout(false);`r`n            this.tabGenerale.PerformLayout();`r`n            this.tabBilancia.ResumeLayout(false);`r`n            this.tabBilancia.PerformLayout();`r`n            this.tabListini.ResumeLayout(false);`r`n            this.tabListini.PerformLayout();"
$content = $content -replace 'this\.pnlArt\.ResumeLayout\(false\);', ($resume + "`r`n            this.pnlArt.ResumeLayout(false);")

# Save as UTF8 with BOM to correctly persist C# files in Visual Studio
[System.IO.File]::WriteAllText("C:\ApProject\APOffice\frmAnaArticolo.Designer.cs", $content, [System.Text.Encoding]::UTF8)
Write-Host "Done"
