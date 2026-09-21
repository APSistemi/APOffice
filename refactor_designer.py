import re
import sys

try:
    with open("frmAnaArticolo.Designer.cs", "r", encoding="utf-8") as f:
        text = f.read()

    orig_len = len(text)
    print(f"Original file length: {orig_len}")

    # Helper function to do safe replacement and report
    def replace_safe(old, new, desc):
        global text
        if old in text:
            text = text.replace(old, new)
            print(f"SUCCESS: {desc}")
        else:
            print(f"FAILED: {desc}")

    # Add TabControl and TabPages to declarations
    decls = """        private System.Windows.Forms.TabControl tabMain;
        private System.Windows.Forms.TabPage tabGenerale;
        private System.Windows.Forms.TabPage tabBilancia;
        private System.Windows.Forms.TabPage tabListini;"""
    replace_safe("private System.Windows.Forms.Panel pnlArt;", "private System.Windows.Forms.Panel pnlArt;\n" + decls, "Declarations")

    # Instantiate
    inst = """            this.tabMain = new System.Windows.Forms.TabControl();
            this.tabGenerale = new System.Windows.Forms.TabPage();
            this.tabBilancia = new System.Windows.Forms.TabPage();
            this.tabListini = new System.Windows.Forms.TabPage();"""
    replace_safe("this.pnlArt = new System.Windows.Forms.Panel();", "this.pnlArt = new System.Windows.Forms.Panel();\n" + inst, "Instantiation")

    # SuspendLayout
    susp = """            this.tabMain.SuspendLayout();
            this.tabGenerale.SuspendLayout();
            this.tabBilancia.SuspendLayout();
            this.tabListini.SuspendLayout();"""
    replace_safe("this.pnlArt.SuspendLayout();", "this.pnlArt.SuspendLayout();\n" + susp, "SuspendLayout")

    # Assign controls from pnlArt to Tabs
    replace_safe("this.pnlArt.Controls.Add(this.panel1);", "this.tabGenerale.Controls.Add(this.panel1);", "panel1 to tabGenerale")
    replace_safe("this.pnlArt.Controls.Add(this.groupBox2);", "this.tabGenerale.Controls.Add(this.groupBox2);", "groupBox2 to tabGenerale")

    replace_safe("this.pnlArt.Controls.Add(this.groupBox1);", "this.tabBilancia.Controls.Add(this.groupBox1);", "groupBox1 to tabBilancia")
    replace_safe("this.pnlArt.Controls.Add(this.groupBox3);", "this.tabBilancia.Controls.Add(this.groupBox3);", "groupBox3 to tabBilancia")

    replace_safe("this.pnlArt.Controls.Add(this.label3);", "this.tabListini.Controls.Add(this.label3);", "label3 to listini")
    replace_safe("this.pnlArt.Controls.Add(this.txtEan);", "this.tabListini.Controls.Add(this.txtEan);", "txtEan to listini")
    replace_safe("this.pnlArt.Controls.Add(this.label38);", "this.tabListini.Controls.Add(this.label38);", "label38 to listini")
    replace_safe("this.pnlArt.Controls.Add(this.btnEanAutomatico);", "this.tabListini.Controls.Add(this.btnEanAutomatico);", "btnEan to listini")
    replace_safe("this.pnlArt.Controls.Add(this.dgv1);", "this.tabListini.Controls.Add(this.dgv1);", "dgv1 to listini")

    replace_safe("this.pnlArt.Controls.Add(this.label1);", "this.tabListini.Controls.Add(this.label1);", "label1 to listini")
    replace_safe("this.pnlArt.Controls.Add(this.dgv2);", "this.tabListini.Controls.Add(this.dgv2);", "dgv2 to listini")
    replace_safe("this.pnlArt.Controls.Add(this.btnForCos);", "this.tabListini.Controls.Add(this.btnForCos);", "btnForCos to listini")

    replace_safe("this.pnlArt.Controls.Add(this.label2);", "this.tabListini.Controls.Add(this.label2);", "label2 to listini")
    replace_safe("this.pnlArt.Controls.Add(this.dgv3);", "this.tabListini.Controls.Add(this.dgv3);", "dgv3 to listini")
    replace_safe("this.pnlArt.Controls.Add(this.label39);", "this.tabListini.Controls.Add(this.label39);", "label39 to listini")
    replace_safe("this.pnlArt.Controls.Add(this.cmbLisVen);", "this.tabListini.Controls.Add(this.cmbLisVen);", "cmbLisVen to listini")

    # Add TabControl to pnlArt
    tab_config = """            // 
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
    """
    replace_safe("this.pnlArt.TabIndex = 1;", "this.pnlArt.TabIndex = 1;\n            this.pnlArt.Controls.Add(this.tabMain);\n" + tab_config, "TabMain Config")

    # Update positions to look good inside tabs using Regex
    def re_sub_safe(pattern, repl, desc):
        global text
        if re.search(pattern, text):
            text = re.sub(pattern, repl, text)
            print(f"SUCCESS (Regex): {desc}")
        else:
            print(f"FAILED (Regex): {desc}")

    re_sub_safe(r"this\.panel1\.Location = new System\.Drawing\.Point\([^\)]+\);", "this.panel1.Location = new System.Drawing.Point(5, 5);", "panel1 Location")
    re_sub_safe(r"this\.groupBox2\.Location = new System\.Drawing\.Point\([^\)]+\);", "this.groupBox2.Location = new System.Drawing.Point(5, 155);", "groupBox2 Location")
    re_sub_safe(r"this\.groupBox1\.Location = new System\.Drawing\.Point\([^\)]+\);", "this.groupBox1.Location = new System.Drawing.Point(10, 10);", "groupBox1 Location")
    re_sub_safe(r"this\.groupBox1\.Size = new System\.Drawing\.Size\([^\)]+\);", "this.groupBox1.Size = new System.Drawing.Size(275, 480);", "groupBox1 Size")
    re_sub_safe(r"this\.groupBox3\.Location = new System\.Drawing\.Point\([^\)]+\);", "this.groupBox3.Location = new System.Drawing.Point(300, 10);", "groupBox3 Location")
    re_sub_safe(r"this\.groupBox3\.Size = new System\.Drawing\.Size\([^\)]+\);", "this.groupBox3.Size = new System.Drawing.Size(330, 480);", "groupBox3 Size")

    # Modernize grids and buttons
    re_sub_safe(r"this\.label3\.Location = new System\.Drawing\.Point\([^\)]+\);", "this.label3.Location = new System.Drawing.Point(10, 15);", "label3 loc")
    re_sub_safe(r"this\.txtEan\.Location = new System\.Drawing\.Point\([^\)]+\);", "this.txtEan.Location = new System.Drawing.Point(70, 12);", "txtEan loc")
    re_sub_safe(r"this\.label38\.Location = new System\.Drawing\.Point\([^\)]+\);", "this.label38.Location = new System.Drawing.Point(205, 15);", "label38 loc")
    re_sub_safe(r"this\.btnEanAutomatico\.Location = new System\.Drawing\.Point\([^\)]+\);", "this.btnEanAutomatico.Location = new System.Drawing.Point(250, 10);", "btnEan loc")
    re_sub_safe(r"this\.dgv1\.Location = new System\.Drawing\.Point\([^\)]+\);", "this.dgv1.Location = new System.Drawing.Point(10, 40);", "dgv1 loc")
    re_sub_safe(r"this\.dgv1\.Size = new System\.Drawing\.Size\([^\)]+\);", "this.dgv1.Size = new System.Drawing.Size(270, 460);", "dgv1 size")

    re_sub_safe(r"this\.label1\.Location = new System\.Drawing\.Point\([^\)]+\);", "this.label1.Location = new System.Drawing.Point(295, 15);", "label1 loc")
    re_sub_safe(r"this\.btnForCos\.Location = new System\.Drawing\.Point\([^\)]+\);", "this.btnForCos.Location = new System.Drawing.Point(450, 10);", "btnForCos loc")
    re_sub_safe(r"this\.dgv2\.Location = new System\.Drawing\.Point\([^\)]+\);", "this.dgv2.Location = new System.Drawing.Point(295, 40);", "dgv2 loc")
    re_sub_safe(r"this\.dgv2\.Size = new System\.Drawing\.Size\([^\)]+\);", "this.dgv2.Size = new System.Drawing.Size(280, 460);", "dgv2 size")

    re_sub_safe(r"this\.label2\.Location = new System\.Drawing\.Point\([^\)]+\);", "this.label2.Location = new System.Drawing.Point(590, 15);", "label2 loc")
    re_sub_safe(r"this\.label39\.Location = new System\.Drawing\.Point\([^\)]+\);", "this.label39.Location = new System.Drawing.Point(590, 15);", "label39 loc") 
    re_sub_safe(r"this\.cmbLisVen\.Location = new System\.Drawing\.Point\([^\)]+\);", "this.cmbLisVen.Location = new System.Drawing.Point(700, 10);", "cmbLisVen loc")
    re_sub_safe(r"this\.dgv3\.Location = new System\.Drawing\.Point\([^\)]+\);", "this.dgv3.Location = new System.Drawing.Point(590, 40);", "dgv3 loc")
    re_sub_safe(r"this\.dgv3\.Size = new System\.Drawing\.Size\([^\)]+\);", "this.dgv3.Size = new System.Drawing.Size(280, 460);", "dgv3 size")

    # Modern colours
    replace_safe("System.Drawing.Color.FromArgb(((int)(((byte)(224)))), ((int)(((byte)(224)))), ((int)(((byte)(224)))))", "System.Drawing.Color.WhiteSmoke", "colors to WhiteSmoke")
    replace_safe("this.panel1.BackColor = System.Drawing.Color.WhiteSmoke;", "this.panel1.BackColor = System.Drawing.Color.White;", "panel1 White")

    # modern form font
    replace_safe('this.Text = " Anagrafica articolo";', 'this.Text = " Anagrafica articolo";\n            this.Font = new System.Drawing.Font("Segoe UI", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));', "Form font")

    # Resume Layouts
    resume = """            this.tabMain.ResumeLayout(false);
            this.tabGenerale.ResumeLayout(false);
            this.tabGenerale.PerformLayout();
            this.tabBilancia.ResumeLayout(false);
            this.tabBilancia.PerformLayout();
            this.tabListini.ResumeLayout(false);
            this.tabListini.PerformLayout();"""
    replace_safe("this.pnlArt.ResumeLayout(false);", resume + "\n            this.pnlArt.ResumeLayout(false);", "ResumeLayouts")

    print(f"Final file length: {len(text)}")
    with open("frmAnaArticolo.Designer.cs", "w", encoding="utf-8") as f:
        f.write(text)
    print("Redesign script completed and wrote to disk", flush=True)
except Exception as e:
    print(f"Error: {e}")
