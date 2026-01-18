using System;
using System.Drawing;
using System.Reflection;
using System.Windows.Forms;
using System.Diagnostics;

namespace QSBar
{
    public partial class HelpForm : Form
    {
        private Label lblTitle;
        private Label lblVersion;
        private Label lblUpdateDate;
        private Label lblShortcuts;
        private LinkLabel lnkDownload;
        private Button btnUpdate;
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
            this.btnUpdate = new Button();
            this.btnClose = new Button();

            // Form
            this.Text = "关于 QS工具箱";
            this.Size = new Size(400, 380);
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.StartPosition = FormStartPosition.CenterScreen;
            this.BackColor = Color.White;

            // Title
            lblTitle.Text = "QS 工具箱 (QSBar)";
            lblTitle.Font = new Font("微软雅黑", 14, FontStyle.Bold);
            lblTitle.Location = new Point(20, 20);
            lblTitle.Size = new Size(360, 30);

            // Version
            lblVersion.Location = new Point(20, 60);
            lblVersion.Size = new Size(360, 20);
            lblVersion.Font = new Font("微软雅黑", 9);

            // Update Date
            lblUpdateDate.Location = new Point(20, 85);
            lblUpdateDate.Size = new Size(360, 20);
            lblUpdateDate.Font = new Font("微软雅黑", 9);

            // Shortcuts
            lblShortcuts.Location = new Point(20, 120);
            lblShortcuts.Size = new Size(360, 120);
            lblShortcuts.Font = new Font("微软雅黑", 9);
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
            lnkDownload.Location = new Point(20, 250);
            lnkDownload.Size = new Size(360, 20);
            lnkDownload.LinkClicked += (s, e) => Process.Start("https://gitee.com/kevin137/qsbar");

            // Buttons
            btnUpdate.Text = "检查更新";
            btnUpdate.Location = new Point(180, 290);
            btnUpdate.Size = new Size(90, 30);
            btnUpdate.Click += async (s, e) => {
                btnUpdate.Enabled = false;
                btnUpdate.Text = "正在检查...";
                await UpdateManager.CheckForUpdateAsync(false);
                btnUpdate.Enabled = true;
                btnUpdate.Text = "检查更新";
            };

            btnClose.Text = "关闭";
            btnClose.Location = new Point(280, 290);
            btnClose.Size = new Size(90, 30);
            btnClose.Click += (s, e) => this.Close();

            this.Controls.AddRange(new Control[] { lblTitle, lblVersion, lblUpdateDate, lblShortcuts, lnkDownload, btnUpdate, btnClose });
        }

        private void LoadInfo()
        {
            var assembly = Assembly.GetExecutingAssembly();
            var version = assembly.GetName().Version;
            var buildDate = new DateTime(2000, 1, 1).AddDays(version.Build).AddSeconds(version.Revision * 2);

            lblVersion.Text = $"版本号: v{version.Major}.{version.Minor}.{version.Build}.{version.Revision}";
            lblUpdateDate.Text = $"最后编译日期: {System.IO.File.GetLastWriteTime(assembly.Location):yyyy-MM-dd HH:mm:ss}";
        }
    }
}
