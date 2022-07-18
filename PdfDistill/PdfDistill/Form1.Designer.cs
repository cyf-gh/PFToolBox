namespace PdfDistill
{
    partial class Form1
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
            this.label1 = new System.Windows.Forms.Label();
            this.btn_GenExcel = new System.Windows.Forms.Button();
            this.lb_Status = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // btn_newPswd
            // 
            this.btn_newPswd.Location = new System.Drawing.Point(12, 12);
            this.btn_newPswd.Name = "btn_newPswd";
            this.btn_newPswd.Size = new System.Drawing.Size(157, 38);
            this.btn_newPswd.TabIndex = 0;
            this.btn_newPswd.Text = "设置新密码";
            this.btn_newPswd.UseVisualStyleBackColor = true;
            this.btn_newPswd.Click += new System.EventHandler(this.btn_newPswd_Click);
            // 
            // lb_pswd
            // 
            this.lb_pswd.AutoSize = true;
            this.lb_pswd.Location = new System.Drawing.Point(243, 22);
            this.lb_pswd.Name = "lb_pswd";
            this.lb_pswd.Size = new System.Drawing.Size(62, 18);
            this.lb_pswd.TabIndex = 1;
            this.lb_pswd.Text = "label1";
            // 
            // btn_Start
            // 
            this.btn_Start.Location = new System.Drawing.Point(12, 400);
            this.btn_Start.Name = "btn_Start";
            this.btn_Start.Size = new System.Drawing.Size(157, 38);
            this.btn_Start.TabIndex = 0;
            this.btn_Start.Text = "开始转化";
            this.btn_Start.UseVisualStyleBackColor = true;
            this.btn_Start.Click += new System.EventHandler(this.btn_Start_Click);
            // 
            // lb_t1
            // 
            this.lb_t1.AutoSize = true;
            this.lb_t1.Location = new System.Drawing.Point(175, 22);
            this.lb_t1.Name = "lb_t1";
            this.lb_t1.Size = new System.Drawing.Size(62, 18);
            this.lb_t1.TabIndex = 2;
            this.lb_t1.Text = "密码：";
            // 
            // btn_selectDir
            // 
            this.btn_selectDir.Location = new System.Drawing.Point(12, 56);
            this.btn_selectDir.Name = "btn_selectDir";
            this.btn_selectDir.Size = new System.Drawing.Size(157, 38);
            this.btn_selectDir.TabIndex = 0;
            this.btn_selectDir.Text = "选取新路径";
            this.btn_selectDir.UseVisualStyleBackColor = true;
            this.btn_selectDir.Click += new System.EventHandler(this.btn_selectDir_Click);
            // 
            // label_tickDir
            // 
            this.label_tickDir.AutoSize = true;
            this.label_tickDir.Location = new System.Drawing.Point(175, 66);
            this.label_tickDir.Name = "label_tickDir";
            this.label_tickDir.Size = new System.Drawing.Size(98, 18);
            this.label_tickDir.TabIndex = 2;
            this.label_tickDir.Text = "票据路径：";
            // 
            // lb_ticketsPath
            // 
            this.lb_ticketsPath.AutoSize = true;
            this.lb_ticketsPath.Location = new System.Drawing.Point(279, 66);
            this.lb_ticketsPath.Name = "lb_ticketsPath";
            this.lb_ticketsPath.Size = new System.Drawing.Size(62, 18);
            this.lb_ticketsPath.TabIndex = 1;
            this.lb_ticketsPath.Text = "label1";
            // 
            // tb_log
            // 
            this.tb_log.AcceptsReturn = true;
            this.tb_log.AcceptsTab = true;
            this.tb_log.Location = new System.Drawing.Point(12, 136);
            this.tb_log.Multiline = true;
            this.tb_log.Name = "tb_log";
            this.tb_log.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.tb_log.Size = new System.Drawing.Size(1008, 258);
            this.tb_log.TabIndex = 3;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(12, 115);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(62, 18);
            this.label1.TabIndex = 2;
            this.label1.Text = "输出：";
            // 
            // btn_GenExcel
            // 
            this.btn_GenExcel.Location = new System.Drawing.Point(178, 400);
            this.btn_GenExcel.Name = "btn_GenExcel";
            this.btn_GenExcel.Size = new System.Drawing.Size(163, 38);
            this.btn_GenExcel.TabIndex = 4;
            this.btn_GenExcel.Text = "生成表格";
            this.btn_GenExcel.UseVisualStyleBackColor = true;
            this.btn_GenExcel.Click += new System.EventHandler(this.btn_GenExcel_Click);
            // 
            // lb_Status
            // 
            this.lb_Status.AutoSize = true;
            this.lb_Status.Location = new System.Drawing.Point(347, 410);
            this.lb_Status.Name = "lb_Status";
            this.lb_Status.Size = new System.Drawing.Size(62, 18);
            this.lb_Status.TabIndex = 5;
            this.lb_Status.Text = "label2";
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(9F, 18F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1032, 450);
            this.Controls.Add(this.lb_Status);
            this.Controls.Add(this.btn_GenExcel);
            this.Controls.Add(this.tb_log);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.label_tickDir);
            this.Controls.Add(this.lb_t1);
            this.Controls.Add(this.lb_ticketsPath);
            this.Controls.Add(this.lb_pswd);
            this.Controls.Add(this.btn_Start);
            this.Controls.Add(this.btn_selectDir);
            this.Controls.Add(this.btn_newPswd);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow;
            this.Name = "Form1";
            this.Text = "电子回单汇总";
            this.Load += new System.EventHandler(this.Form1_Load);
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
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button btn_GenExcel;
        private System.Windows.Forms.Label lb_Status;
    }
}

