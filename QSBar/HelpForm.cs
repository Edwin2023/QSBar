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
        private TextBox txtContent; // 使用 TextBox 替代 Label 以支持滚动条
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
            this.txtContent = new TextBox();
            this.lnkDownload = new LinkLabel();
            this.btnClose = new Button();

            // Form
            this.Text = "关于 QS 工具箱";
            this.Size = new Size(550, 650); // 进一步扩大窗口
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.StartPosition = FormStartPosition.CenterScreen;
            this.BackColor = Color.White;
            this.Font = new Font("微软雅黑", 10);
            this.KeyPreview = true; // 允许窗体接收键盘事件

            // ESC 键关闭
            this.KeyDown += (s, e) => {
                if (e.KeyCode == Keys.Escape) this.Close();
            };

            // Title
            lblTitle.Text = "QS 工具箱 (QSBar)";
            lblTitle.Font = new Font("微软雅黑", 18, FontStyle.Bold);
            lblTitle.Location = new Point(20, 20);
            lblTitle.Size = new Size(500, 45);
            lblTitle.ForeColor = Color.FromArgb(0, 120, 215);

            // Version
            lblVersion.Location = new Point(25, 75);
            lblVersion.Size = new Size(500, 30);

            // Update Date
            lblUpdateDate.Location = new Point(25, 105);
            lblUpdateDate.Size = new Size(500, 30);

            // Content TextBox (Shortcuts & Features)
            txtContent.Multiline = true;
            txtContent.ReadOnly = true;
            txtContent.ScrollBars = ScrollBars.Vertical;
            txtContent.BorderStyle = BorderStyle.None;
            txtContent.BackColor = Color.White;
            txtContent.Location = new Point(25, 145);
            txtContent.Size = new Size(485, 380); // 较大的显示区域
            txtContent.Text = "常用快捷键:\r\n" +
                               "--------------------------------------------------\r\n" +
                               "Ctrl + 5:  选中可见单元格\r\n" +
                               "Ctrl + 6:  亿/万位格式切换\r\n" +
                               "Ctrl + 7:  文本换行格式\r\n" +
                               "Ctrl + 8:  会计格式 (0位小数)\r\n" +
                               "Ctrl + 9:  会计格式 (2位小数)\r\n" +
                               "Ctrl + 0:  会计格式 (3位小数)\r\n\r\n" +
                               "功能说明:\r\n" +
                               "--------------------------------------------------\r\n" +
                               "【批量处理】支持大数据量的快速计算与转换。\r\n" +
                               "【导出报表】支持标准、内部等多种报表导出格式。\r\n" +
                               "【格式规范】一键数值化、公式锁定、换行处理。\r\n" +
                               "【分级筛选】自动设置分级显示，美化分级样式。\r\n" +
                               "【图片工具】批量调整图片大小，全选当前页图片。\r\n" +
                               "【表/文件管理】合并工作表、生成目录、文件管理。\r\n" +
                               "【计算模式】快速切换自动/手动计算模式。";

            // Download Link
            lnkDownload.Text = "下载地址: https://gitee.com/kevin137/qsbar";
            lnkDownload.Location = new Point(25, 540);
            lnkDownload.Size = new Size(500, 30);
            lnkDownload.LinkClicked += (s, e) => {
                try { Process.Start("https://gitee.com/kevin137/qsbar"); }
                catch { MessageBox.Show("无法打开链接。"); }
            };

            // Buttons
            btnClose.Text = "关闭";
            btnClose.Location = new Point(410, 565);
            btnClose.Size = new Size(100, 35);
            btnClose.Click += (s, e) => this.Close();

            this.Controls.AddRange(new Control[] { lblTitle, lblVersion, lblUpdateDate, txtContent, lnkDownload, btnClose });
        }

        private void LoadInfo()
        {
            var assembly = Assembly.GetExecutingAssembly();
            var version = assembly.GetName().Version;
            
            lblVersion.Text = string.Format("版本号: v{0}.{1}.{2}.{3}", version.Major, version.Minor, version.Build, version.Revision);
            
            try {
                var lastWriteTime = File.GetLastWriteTime(assembly.Location);
                lblUpdateDate.Text = string.Format("最后编译日期: {0:yyyy-MM-dd HH:mm:ss}", lastWriteTime);
            }
            catch {
                lblUpdateDate.Text = "更新日期: 无法获取";
            }
        }
    }
}
