using System;
using System.Drawing;
using System.Reflection;
using System.Windows.Forms;
using System.Diagnostics;
using System.IO;

namespace QSBar
{
    public partial class HelpForm : Form
    {
        private Label lblTitle;
        private Label lblVersion;
        private Label lblUpdateDate;
        private Label lblShortcuts;
        private LinkLabel lnkDownload;
        private Button btnClose;

        public HelpForm()
        {
            InitializeComponent();
            LoadInfo();
        }

        private void InitializeComponent()
        {
            this.lblTitle = new Label();
            this.lblVersion = new Label();
            this.lblUpdateDate = new Label();
            this.lblShortcuts = new Label();
            this.lnkDownload = new LinkLabel();
            this.btnClose = new Button();

            // Form
            this.Text = "关于 QS 工具箱";
            this.Size = new Size(480, 500);
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.StartPosition = FormStartPosition.CenterScreen;
            this.BackColor = Color.White;
            this.Font = new Font("微软雅黑", 10);

            // Title
            lblTitle.Text = "QS 工具箱 (QSBar)";
            lblTitle.Font = new Font("微软雅黑", 16, FontStyle.Bold);
            lblTitle.Location = new Point(20, 20);
            lblTitle.Size = new Size(440, 40);
            lblTitle.ForeColor = Color.FromArgb(0, 120, 215);

            // Version
            lblVersion.Location = new Point(25, 75);
            lblVersion.Size = new Size(420, 30);

            // Update Date
            lblUpdateDate.Location = new Point(25, 105);
            lblUpdateDate.Size = new Size(420, 30);

            // Shortcuts
            lblShortcuts.Location = new Point(25, 145);
            lblShortcuts.Size = new Size(420, 200);
            lblShortcuts.Text = "常用快捷键:\n" +
                               "------------------------------------------\n" +
                               "Ctrl + 5:  选中可见单元格\n" +
                               "Ctrl + 6:  亿/万位格式切换\n" +
                               "Ctrl + 7:  文本换行格式\n" +
                               "Ctrl + 8:  会计格式 (0位小数)\n" +
                               "Ctrl + 9:  会计格式 (2位小数)\n" +
                               "Ctrl + 0:  会计格式 (3位小数)";

            // Download Link
            lnkDownload.Text = "下载地址: https://gitee.com/kevin137/qsbar";
            lnkDownload.Location = new Point(25, 360);
            lnkDownload.Size = new Size(420, 30);
            lnkDownload.LinkClicked += (s, e) => {
                try { Process.Start("https://gitee.com/kevin137/qsbar"); }
                catch { MessageBox.Show("无法打开链接。"); }
            };

            // Buttons
            btnClose.Text = "关闭";
            btnClose.Location = new Point(350, 410);
            btnClose.Size = new Size(90, 35);
            btnClose.Click += (s, e) => this.Close();

            this.Controls.AddRange(new Control[] { lblTitle, lblVersion, lblUpdateDate, lblShortcuts, lnkDownload, btnClose });
        }

        private void LoadInfo()
        {
            var assembly = Assembly.GetExecutingAssembly();
            var version = assembly.GetName().Version;
            
            lblVersion.Text = $"版本号: v{version.Major}.{version.Minor}.{version.Build}.{version.Revision}";
            
            try {
                var lastWriteTime = File.GetLastWriteTime(assembly.Location);
                lblUpdateDate.Text = $"最后编译日期: {lastWriteTime:yyyy-MM-dd HH:mm:ss}";
            }
            catch {
                lblUpdateDate.Text = "更新日期: 无法获取";
            }
        }
    }
}
