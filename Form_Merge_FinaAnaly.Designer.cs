namespace MergeExcel {
    partial class Form_Merge_FinaAnaly {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose( bool disposing ) {
            if ( disposing && ( components != null ) ) {
                components.Dispose();
            }
            base.Dispose( disposing );
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent() {
            this.menuStrip1 = new System.Windows.Forms.MenuStrip();
            this.打开文件夹ToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.tscb_SheetSelect = new System.Windows.Forms.ToolStripComboBox();
            this.tabPage3 = new System.Windows.Forms.TabPage();
            this.tv_xjllb = new System.Windows.Forms.TreeView();
            this.tabPage2 = new System.Windows.Forms.TabPage();
            this.tv_lr = new System.Windows.Forms.TreeView();
            this.tabPage1 = new System.Windows.Forms.TabPage();
            this.tv_zcfz = new System.Windows.Forms.TreeView();
            this.tabControl1 = new System.Windows.Forms.TabControl();
            this.menuStrip1.SuspendLayout();
            this.tabPage3.SuspendLayout();
            this.tabPage2.SuspendLayout();
            this.tabPage1.SuspendLayout();
            this.tabControl1.SuspendLayout();
            this.SuspendLayout();
            // 
            // menuStrip1
            // 
            this.menuStrip1.ImageScalingSize = new System.Drawing.Size(20, 20);
            this.menuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.打开文件夹ToolStripMenuItem,
            this.tscb_SheetSelect});
            this.menuStrip1.Location = new System.Drawing.Point(0, 0);
            this.menuStrip1.Name = "menuStrip1";
            this.menuStrip1.Size = new System.Drawing.Size(944, 32);
            this.menuStrip1.TabIndex = 0;
            this.menuStrip1.Text = "menuStrip1";
            // 
            // 打开文件夹ToolStripMenuItem
            // 
            this.打开文件夹ToolStripMenuItem.Name = "打开文件夹ToolStripMenuItem";
            this.打开文件夹ToolStripMenuItem.Size = new System.Drawing.Size(98, 28);
            this.打开文件夹ToolStripMenuItem.Text = "打开文件夹";
            // 
            // tscb_SheetSelect
            // 
            this.tscb_SheetSelect.Name = "tscb_SheetSelect";
            this.tscb_SheetSelect.Size = new System.Drawing.Size(121, 28);
            // 
            // tabPage3
            // 
            this.tabPage3.Controls.Add(this.tv_xjllb);
            this.tabPage3.Location = new System.Drawing.Point(4, 25);
            this.tabPage3.Name = "tabPage3";
            this.tabPage3.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage3.Size = new System.Drawing.Size(936, 446);
            this.tabPage3.TabIndex = 2;
            this.tabPage3.Text = "现金流量表";
            this.tabPage3.UseVisualStyleBackColor = true;
            // 
            // tv_xjllb
            // 
            this.tv_xjllb.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tv_xjllb.Location = new System.Drawing.Point(3, 3);
            this.tv_xjllb.Name = "tv_xjllb";
            this.tv_xjllb.Size = new System.Drawing.Size(930, 440);
            this.tv_xjllb.TabIndex = 5;
            // 
            // tabPage2
            // 
            this.tabPage2.Controls.Add(this.tv_lr);
            this.tabPage2.Location = new System.Drawing.Point(4, 25);
            this.tabPage2.Name = "tabPage2";
            this.tabPage2.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage2.Size = new System.Drawing.Size(936, 446);
            this.tabPage2.TabIndex = 1;
            this.tabPage2.Text = "利润表";
            this.tabPage2.UseVisualStyleBackColor = true;
            // 
            // tv_lr
            // 
            this.tv_lr.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tv_lr.Location = new System.Drawing.Point(3, 3);
            this.tv_lr.Name = "tv_lr";
            this.tv_lr.Size = new System.Drawing.Size(930, 440);
            this.tv_lr.TabIndex = 4;
            // 
            // tabPage1
            // 
            this.tabPage1.Controls.Add(this.tv_zcfz);
            this.tabPage1.Location = new System.Drawing.Point(4, 25);
            this.tabPage1.Name = "tabPage1";
            this.tabPage1.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage1.Size = new System.Drawing.Size(936, 446);
            this.tabPage1.TabIndex = 0;
            this.tabPage1.Text = "资产负债表";
            this.tabPage1.UseVisualStyleBackColor = true;
            // 
            // tv_zcfz
            // 
            this.tv_zcfz.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tv_zcfz.Location = new System.Drawing.Point(3, 3);
            this.tv_zcfz.Name = "tv_zcfz";
            this.tv_zcfz.Size = new System.Drawing.Size(930, 440);
            this.tv_zcfz.TabIndex = 3;
            // 
            // tabControl1
            // 
            this.tabControl1.Controls.Add(this.tabPage1);
            this.tabControl1.Controls.Add(this.tabPage2);
            this.tabControl1.Controls.Add(this.tabPage3);
            this.tabControl1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tabControl1.Location = new System.Drawing.Point(0, 32);
            this.tabControl1.Name = "tabControl1";
            this.tabControl1.SelectedIndex = 0;
            this.tabControl1.Size = new System.Drawing.Size(944, 475);
            this.tabControl1.TabIndex = 3;
            // 
            // Form_Merge_FinaAnaly
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(944, 507);
            this.Controls.Add(this.tabControl1);
            this.Controls.Add(this.menuStrip1);
            this.MainMenuStrip = this.menuStrip1;
            this.Name = "Form_Merge_FinaAnaly";
            this.Text = "Form_Merge_FinaAnaly";
            this.Load += new System.EventHandler(this.Form_Merge_FinaAnaly_Load);
            this.menuStrip1.ResumeLayout(false);
            this.menuStrip1.PerformLayout();
            this.tabPage3.ResumeLayout(false);
            this.tabPage2.ResumeLayout(false);
            this.tabPage1.ResumeLayout(false);
            this.tabControl1.ResumeLayout(false);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.MenuStrip menuStrip1;
        private System.Windows.Forms.ToolStripMenuItem 打开文件夹ToolStripMenuItem;
        private System.Windows.Forms.ToolStripComboBox tscb_SheetSelect;
        private System.Windows.Forms.TabPage tabPage3;
        private System.Windows.Forms.TreeView tv_xjllb;
        private System.Windows.Forms.TabPage tabPage2;
        private System.Windows.Forms.TreeView tv_lr;
        private System.Windows.Forms.TabPage tabPage1;
        private System.Windows.Forms.TreeView tv_zcfz;
        private System.Windows.Forms.TabControl tabControl1;
    }
}