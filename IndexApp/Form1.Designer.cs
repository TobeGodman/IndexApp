namespace IndexApp
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
        /// 设计器支持所需的方法 - 不要
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form1));
            this.connectBt = new System.Windows.Forms.Button();
            this.getIdBt = new System.Windows.Forms.Button();
            this.mainTxtBox = new System.Windows.Forms.TextBox();
            this.exportBt = new System.Windows.Forms.Button();
            this.tabControl1 = new System.Windows.Forms.TabControl();
            this.folderSelectBt = new System.Windows.Forms.Button();
            this.excelPathTxtBox = new System.Windows.Forms.TextBox();
            this.tabControl2 = new System.Windows.Forms.TabControl();
            this.settingTabPage = new System.Windows.Forms.TabPage();
            this.infoTabPage = new System.Windows.Forms.TabPage();
            this.tabControl2.SuspendLayout();
            this.infoTabPage.SuspendLayout();
            this.SuspendLayout();
            // 
            // connectBt
            // 
            this.connectBt.Location = new System.Drawing.Point(29, 16);
            this.connectBt.Margin = new System.Windows.Forms.Padding(2, 3, 2, 3);
            this.connectBt.Name = "connectBt";
            this.connectBt.Size = new System.Drawing.Size(100, 33);
            this.connectBt.TabIndex = 0;
            this.connectBt.Text = "连接服务器";
            this.connectBt.UseVisualStyleBackColor = true;
            this.connectBt.Click += new System.EventHandler(this.connectBt_Click);
            // 
            // getIdBt
            // 
            this.getIdBt.Location = new System.Drawing.Point(145, 16);
            this.getIdBt.Margin = new System.Windows.Forms.Padding(2, 3, 2, 3);
            this.getIdBt.Name = "getIdBt";
            this.getIdBt.Size = new System.Drawing.Size(100, 33);
            this.getIdBt.TabIndex = 2;
            this.getIdBt.Text = "获取合约信息";
            this.getIdBt.UseVisualStyleBackColor = true;
            this.getIdBt.Click += new System.EventHandler(this.getIdBt_Click);
            // 
            // mainTxtBox
            // 
            this.mainTxtBox.Location = new System.Drawing.Point(0, 0);
            this.mainTxtBox.Margin = new System.Windows.Forms.Padding(2, 3, 2, 3);
            this.mainTxtBox.Multiline = true;
            this.mainTxtBox.Name = "mainTxtBox";
            this.mainTxtBox.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.mainTxtBox.Size = new System.Drawing.Size(234, 432);
            this.mainTxtBox.TabIndex = 4;
            // 
            // exportBt
            // 
            this.exportBt.Location = new System.Drawing.Point(372, 16);
            this.exportBt.Margin = new System.Windows.Forms.Padding(2, 3, 2, 3);
            this.exportBt.Name = "exportBt";
            this.exportBt.Size = new System.Drawing.Size(100, 33);
            this.exportBt.TabIndex = 7;
            this.exportBt.Text = "导出到EXCEL";
            this.exportBt.UseVisualStyleBackColor = true;
            this.exportBt.Click += new System.EventHandler(this.exportBt_Click);
            // 
            // tabControl1
            // 
            this.tabControl1.Location = new System.Drawing.Point(29, 55);
            this.tabControl1.Margin = new System.Windows.Forms.Padding(2, 3, 2, 3);
            this.tabControl1.Name = "tabControl1";
            this.tabControl1.SelectedIndex = 0;
            this.tabControl1.Size = new System.Drawing.Size(521, 462);
            this.tabControl1.TabIndex = 9;
            // 
            // folderSelectBt
            // 
            this.folderSelectBt.Location = new System.Drawing.Point(477, 16);
            this.folderSelectBt.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.folderSelectBt.Name = "folderSelectBt";
            this.folderSelectBt.Size = new System.Drawing.Size(73, 33);
            this.folderSelectBt.TabIndex = 10;
            this.folderSelectBt.Text = "文件夹:";
            this.folderSelectBt.UseVisualStyleBackColor = true;
            this.folderSelectBt.Click += new System.EventHandler(this.folderSelectBt_Click);
            // 
            // excelPathTxtBox
            // 
            this.excelPathTxtBox.Location = new System.Drawing.Point(556, 21);
            this.excelPathTxtBox.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.excelPathTxtBox.Name = "excelPathTxtBox";
            this.excelPathTxtBox.Size = new System.Drawing.Size(224, 23);
            this.excelPathTxtBox.TabIndex = 11;
            this.excelPathTxtBox.Text = "D:\\";
            // 
            // tabControl2
            // 
            this.tabControl2.Controls.Add(this.settingTabPage);
            this.tabControl2.Controls.Add(this.infoTabPage);
            this.tabControl2.Location = new System.Drawing.Point(556, 55);
            this.tabControl2.Name = "tabControl2";
            this.tabControl2.SelectedIndex = 0;
            this.tabControl2.Size = new System.Drawing.Size(242, 462);
            this.tabControl2.TabIndex = 12;
            // 
            // settingTabPage
            // 
            this.settingTabPage.Location = new System.Drawing.Point(4, 26);
            this.settingTabPage.Name = "settingTabPage";
            this.settingTabPage.Padding = new System.Windows.Forms.Padding(3);
            this.settingTabPage.Size = new System.Drawing.Size(234, 432);
            this.settingTabPage.TabIndex = 0;
            this.settingTabPage.Text = "保证金设置";
            this.settingTabPage.UseVisualStyleBackColor = true;
            // 
            // infoTabPage
            // 
            this.infoTabPage.Controls.Add(this.mainTxtBox);
            this.infoTabPage.Location = new System.Drawing.Point(4, 26);
            this.infoTabPage.Name = "infoTabPage";
            this.infoTabPage.Padding = new System.Windows.Forms.Padding(3);
            this.infoTabPage.Size = new System.Drawing.Size(234, 432);
            this.infoTabPage.TabIndex = 1;
            this.infoTabPage.Text = "信息";
            this.infoTabPage.UseVisualStyleBackColor = true;
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 17F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.SystemColors.Window;
            this.ClientSize = new System.Drawing.Size(810, 545);
            this.Controls.Add(this.tabControl2);
            this.Controls.Add(this.excelPathTxtBox);
            this.Controls.Add(this.folderSelectBt);
            this.Controls.Add(this.tabControl1);
            this.Controls.Add(this.exportBt);
            this.Controls.Add(this.getIdBt);
            this.Controls.Add(this.connectBt);
            this.Font = new System.Drawing.Font("微软雅黑", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.ForeColor = System.Drawing.SystemColors.WindowText;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Margin = new System.Windows.Forms.Padding(2, 3, 2, 3);
            this.Name = "Form1";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "期货指数";
            this.Load += new System.EventHandler(this.Form1_Load);
            this.tabControl2.ResumeLayout(false);
            this.infoTabPage.ResumeLayout(false);
            this.infoTabPage.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button connectBt;
        private System.Windows.Forms.Button getIdBt;
        private System.Windows.Forms.TextBox mainTxtBox;
        private System.Windows.Forms.Button exportBt;
        private System.Windows.Forms.TabControl tabControl1;
        private System.Windows.Forms.Button folderSelectBt;
        private System.Windows.Forms.TextBox excelPathTxtBox;
        private System.Windows.Forms.TabControl tabControl2;
        private System.Windows.Forms.TabPage settingTabPage;
        private System.Windows.Forms.TabPage infoTabPage;
    }
}

