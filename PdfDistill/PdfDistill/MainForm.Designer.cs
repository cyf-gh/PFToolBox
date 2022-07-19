namespace PdfDistill
{
    partial class MainForm
    {
        /// <summary>
        /// 必需的设计器变量。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// 清理所有正在使用的资源。
        /// </summary>
        /// <param name="disposing">如果应释放托管资源，为 true；否则为 false。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows 窗体设计器生成的代码

        /// <summary>
        /// 设计器支持所需的方法 - 不要修改
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InitializeComponent()
        {
            this.btn_newPswd = new System.Windows.Forms.Button();
            this.lb_pswd = new System.Windows.Forms.Label();
            this.btn_Start = new System.Windows.Forms.Button();
            this.lb_t1 = new System.Windows.Forms.Label();
            this.btn_selectDir = new System.Windows.Forms.Button();
            this.label_tickDir = new System.Windows.Forms.Label();
            this.lb_ticketsPath = new System.Windows.Forms.Label();
            this.tb_log = new System.Windows.Forms.TextBox();
            this.btn_GenExcel = new System.Windows.Forms.Button();
            this.btn_manual = new System.Windows.Forms.Button();
            this.bgw_ExcelGenerate = new System.ComponentModel.BackgroundWorker();
            this.tabControl = new System.Windows.Forms.TabControl();
            this.tabPage1 = new System.Windows.Forms.TabPage();
            this.btn_DeleteItem = new System.Windows.Forms.Button();
            this.lbox_Files = new System.Windows.Forms.ListBox();
            this.tabPage_Output = new System.Windows.Forms.TabPage();
            this.statusStrip1 = new System.Windows.Forms.StatusStrip();
            this.lb_Status = new System.Windows.Forms.ToolStripStatusLabel();
            this.progressBar_Worker = new System.Windows.Forms.ToolStripProgressBar();
            this.btn_OpenExcelPath = new System.Windows.Forms.Button();
            this.tabControl.SuspendLayout();
            this.tabPage1.SuspendLayout();
            this.tabPage_Output.SuspendLayout();
            this.statusStrip1.SuspendLayout();
            this.SuspendLayout();
            // 
            // btn_newPswd
            // 
            this.btn_newPswd.Location = new System.Drawing.Point(11, 10);
            this.btn_newPswd.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.btn_newPswd.Name = "btn_newPswd";
            this.btn_newPswd.Size = new System.Drawing.Size(140, 32);
            this.btn_newPswd.TabIndex = 0;
            this.btn_newPswd.Text = "设置新密码";
            this.btn_newPswd.UseVisualStyleBackColor = true;
            this.btn_newPswd.Click += new System.EventHandler(this.btn_newPswd_Click);
            // 
            // lb_pswd
            // 
            this.lb_pswd.AutoSize = true;
            this.lb_pswd.Location = new System.Drawing.Point(248, 19);
            this.lb_pswd.Name = "lb_pswd";
            this.lb_pswd.Size = new System.Drawing.Size(55, 15);
            this.lb_pswd.TabIndex = 1;
            this.lb_pswd.Text = "label1";
            // 
            // btn_Start
            // 
            this.btn_Start.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(192)))), ((int)(((byte)(255)))), ((int)(((byte)(192)))));
            this.btn_Start.Location = new System.Drawing.Point(11, 333);
            this.btn_Start.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.btn_Start.Name = "btn_Start";
            this.btn_Start.Size = new System.Drawing.Size(140, 32);
            this.btn_Start.TabIndex = 0;
            this.btn_Start.Text = "开始转化";
            this.btn_Start.UseVisualStyleBackColor = false;
            this.btn_Start.Click += new System.EventHandler(this.btn_Start_Click);
            // 
            // lb_t1
            // 
            this.lb_t1.AutoSize = true;
            this.lb_t1.Location = new System.Drawing.Point(156, 18);
            this.lb_t1.Name = "lb_t1";
            this.lb_t1.Size = new System.Drawing.Size(76, 15);
            this.lb_t1.TabIndex = 2;
            this.lb_t1.Text = "PDF密码：";
            // 
            // btn_selectDir
            // 
            this.btn_selectDir.Location = new System.Drawing.Point(11, 47);
            this.btn_selectDir.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.btn_selectDir.Name = "btn_selectDir";
            this.btn_selectDir.Size = new System.Drawing.Size(140, 32);
            this.btn_selectDir.TabIndex = 0;
            this.btn_selectDir.Text = "选取新路径";
            this.btn_selectDir.UseVisualStyleBackColor = true;
            this.btn_selectDir.Click += new System.EventHandler(this.btn_selectDir_Click);
            // 
            // label_tickDir
            // 
            this.label_tickDir.AutoSize = true;
            this.label_tickDir.Location = new System.Drawing.Point(156, 55);
            this.label_tickDir.Name = "label_tickDir";
            this.label_tickDir.Size = new System.Drawing.Size(82, 15);
            this.label_tickDir.TabIndex = 2;
            this.label_tickDir.Text = "票据路径：";
            // 
            // lb_ticketsPath
            // 
            this.lb_ticketsPath.AutoSize = true;
            this.lb_ticketsPath.Location = new System.Drawing.Point(248, 55);
            this.lb_ticketsPath.Name = "lb_ticketsPath";
            this.lb_ticketsPath.Size = new System.Drawing.Size(55, 15);
            this.lb_ticketsPath.TabIndex = 1;
            this.lb_ticketsPath.Text = "label1";
            // 
            // tb_log
            // 
            this.tb_log.AcceptsReturn = true;
            this.tb_log.AcceptsTab = true;
            this.tb_log.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tb_log.Location = new System.Drawing.Point(3, 3);
            this.tb_log.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.tb_log.Multiline = true;
            this.tb_log.Name = "tb_log";
            this.tb_log.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.tb_log.Size = new System.Drawing.Size(880, 209);
            this.tb_log.TabIndex = 3;
            // 
            // btn_GenExcel
            // 
            this.btn_GenExcel.Location = new System.Drawing.Point(158, 333);
            this.btn_GenExcel.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.btn_GenExcel.Name = "btn_GenExcel";
            this.btn_GenExcel.Size = new System.Drawing.Size(145, 32);
            this.btn_GenExcel.TabIndex = 4;
            this.btn_GenExcel.Text = "生成表格";
            this.btn_GenExcel.UseVisualStyleBackColor = true;
            this.btn_GenExcel.Click += new System.EventHandler(this.btn_GenExcel_Click);
            // 
            // btn_manual
            // 
            this.btn_manual.Location = new System.Drawing.Point(783, 333);
            this.btn_manual.Name = "btn_manual";
            this.btn_manual.Size = new System.Drawing.Size(115, 32);
            this.btn_manual.TabIndex = 6;
            this.btn_manual.Text = "软件说明";
            this.btn_manual.UseVisualStyleBackColor = true;
            this.btn_manual.Click += new System.EventHandler(this.btn_manual_Click);
            // 
            // bgw_ExcelGenerate
            // 
            this.bgw_ExcelGenerate.DoWork += new System.ComponentModel.DoWorkEventHandler(this.bgw_ExcelGenerate_DoWork);
            this.bgw_ExcelGenerate.ProgressChanged += new System.ComponentModel.ProgressChangedEventHandler(this.bgw_ExcelGenerate_ProgressChanged);
            this.bgw_ExcelGenerate.RunWorkerCompleted += new System.ComponentModel.RunWorkerCompletedEventHandler(this.bgw_ExcelGenerate_RunWorkerCompleted);
            // 
            // tabControl
            // 
            this.tabControl.Controls.Add(this.tabPage1);
            this.tabControl.Controls.Add(this.tabPage_Output);
            this.tabControl.Location = new System.Drawing.Point(11, 84);
            this.tabControl.Name = "tabControl";
            this.tabControl.SelectedIndex = 0;
            this.tabControl.Size = new System.Drawing.Size(894, 244);
            this.tabControl.TabIndex = 9;
            // 
            // tabPage1
            // 
            this.tabPage1.AllowDrop = true;
            this.tabPage1.Controls.Add(this.btn_DeleteItem);
            this.tabPage1.Controls.Add(this.lbox_Files);
            this.tabPage1.Location = new System.Drawing.Point(4, 25);
            this.tabPage1.Name = "tabPage1";
            this.tabPage1.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage1.Size = new System.Drawing.Size(886, 215);
            this.tabPage1.TabIndex = 0;
            this.tabPage1.Text = "PDF文件";
            this.tabPage1.UseVisualStyleBackColor = true;
            // 
            // btn_DeleteItem
            // 
            this.btn_DeleteItem.Location = new System.Drawing.Point(700, 169);
            this.btn_DeleteItem.Name = "btn_DeleteItem";
            this.btn_DeleteItem.Size = new System.Drawing.Size(161, 29);
            this.btn_DeleteItem.TabIndex = 1;
            this.btn_DeleteItem.Text = "去除选定文件";
            this.btn_DeleteItem.UseVisualStyleBackColor = true;
            this.btn_DeleteItem.Click += new System.EventHandler(this.btn_DeleteItem_Click);
            // 
            // lbox_Files
            // 
            this.lbox_Files.AllowDrop = true;
            this.lbox_Files.Dock = System.Windows.Forms.DockStyle.Fill;
            this.lbox_Files.FormattingEnabled = true;
            this.lbox_Files.ItemHeight = 15;
            this.lbox_Files.Location = new System.Drawing.Point(3, 3);
            this.lbox_Files.Name = "lbox_Files";
            this.lbox_Files.Size = new System.Drawing.Size(880, 209);
            this.lbox_Files.TabIndex = 0;
            this.lbox_Files.DragDrop += new System.Windows.Forms.DragEventHandler(this.lbox_Files_DragDrop);
            this.lbox_Files.DragEnter += new System.Windows.Forms.DragEventHandler(this.lbox_Files_DragEnter);
            this.lbox_Files.DragOver += new System.Windows.Forms.DragEventHandler(this.lbox_Files_DragOver);
            // 
            // tabPage_Output
            // 
            this.tabPage_Output.Controls.Add(this.tb_log);
            this.tabPage_Output.Location = new System.Drawing.Point(4, 25);
            this.tabPage_Output.Name = "tabPage_Output";
            this.tabPage_Output.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage_Output.Size = new System.Drawing.Size(886, 215);
            this.tabPage_Output.TabIndex = 1;
            this.tabPage_Output.Text = "输出";
            this.tabPage_Output.UseVisualStyleBackColor = true;
            // 
            // statusStrip1
            // 
            this.statusStrip1.ImageScalingSize = new System.Drawing.Size(20, 20);
            this.statusStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.progressBar_Worker,
            this.lb_Status});
            this.statusStrip1.Location = new System.Drawing.Point(0, 379);
            this.statusStrip1.Name = "statusStrip1";
            this.statusStrip1.Size = new System.Drawing.Size(917, 26);
            this.statusStrip1.TabIndex = 10;
            this.statusStrip1.Text = "statusStrip1";
            // 
            // lb_Status
            // 
            this.lb_Status.Name = "lb_Status";
            this.lb_Status.Size = new System.Drawing.Size(167, 20);
            this.lb_Status.Text = "toolStripStatusLabel1";
            // 
            // progressBar_Worker
            // 
            this.progressBar_Worker.Name = "progressBar_Worker";
            this.progressBar_Worker.Size = new System.Drawing.Size(120, 18);
            // 
            // btn_OpenExcelPath
            // 
            this.btn_OpenExcelPath.Location = new System.Drawing.Point(309, 334);
            this.btn_OpenExcelPath.Name = "btn_OpenExcelPath";
            this.btn_OpenExcelPath.Size = new System.Drawing.Size(155, 31);
            this.btn_OpenExcelPath.TabIndex = 11;
            this.btn_OpenExcelPath.Text = "打开表格文件夹";
            this.btn_OpenExcelPath.UseVisualStyleBackColor = true;
            this.btn_OpenExcelPath.Click += new System.EventHandler(this.btn_OpenExcelPath_Click);
            // 
            // MainForm
            // 
            this.AllowDrop = true;
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(917, 405);
            this.Controls.Add(this.btn_OpenExcelPath);
            this.Controls.Add(this.statusStrip1);
            this.Controls.Add(this.tabControl);
            this.Controls.Add(this.btn_manual);
            this.Controls.Add(this.btn_GenExcel);
            this.Controls.Add(this.label_tickDir);
            this.Controls.Add(this.lb_t1);
            this.Controls.Add(this.lb_ticketsPath);
            this.Controls.Add(this.lb_pswd);
            this.Controls.Add(this.btn_Start);
            this.Controls.Add(this.btn_selectDir);
            this.Controls.Add(this.btn_newPswd);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow;
            this.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.Name = "MainForm";
            this.Text = "电子回单汇总";
            this.Load += new System.EventHandler(this.Form1_Load);
            this.DragDrop += new System.Windows.Forms.DragEventHandler(this.MainForm_DragDrop);
            this.DragEnter += new System.Windows.Forms.DragEventHandler(this.MainForm_DragEnter);
            this.DragOver += new System.Windows.Forms.DragEventHandler(this.MainForm_DragOver);
            this.tabControl.ResumeLayout(false);
            this.tabPage1.ResumeLayout(false);
            this.tabPage_Output.ResumeLayout(false);
            this.tabPage_Output.PerformLayout();
            this.statusStrip1.ResumeLayout(false);
            this.statusStrip1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btn_newPswd;
        private System.Windows.Forms.Label lb_pswd;
        private System.Windows.Forms.Button btn_Start;
        private System.Windows.Forms.Label lb_t1;
        private System.Windows.Forms.Button btn_selectDir;
        private System.Windows.Forms.Label label_tickDir;
        private System.Windows.Forms.Label lb_ticketsPath;
        private System.Windows.Forms.TextBox tb_log;
        private System.Windows.Forms.Button btn_GenExcel;
        private System.Windows.Forms.Button btn_manual;
        private System.ComponentModel.BackgroundWorker bgw_ExcelGenerate;
        private System.Windows.Forms.TabControl tabControl;
        private System.Windows.Forms.TabPage tabPage1;
        private System.Windows.Forms.TabPage tabPage_Output;
        private System.Windows.Forms.ListBox lbox_Files;
        private System.Windows.Forms.Button btn_DeleteItem;
        private System.Windows.Forms.StatusStrip statusStrip1;
        private System.Windows.Forms.ToolStripStatusLabel lb_Status;
        private System.Windows.Forms.ToolStripProgressBar progressBar_Worker;
        private System.Windows.Forms.Button btn_OpenExcelPath;
    }
}

