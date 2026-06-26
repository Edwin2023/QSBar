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
        private TextBox txtContent;
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

            this.Text = WpsExcelAddIn.UseChineseRibbon ? "关于 QS 工具箱" : "About QS Toolbox";
            this.Size = new Size(650, 860);
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.StartPosition = FormStartPosition.CenterScreen;
            this.BackColor = Color.White;
            this.Font = new Font("Microsoft YaHei", 10);
            this.KeyPreview = true;

            this.KeyDown += (s, e) => {
                if (e.KeyCode == Keys.Escape) this.Close();
            };

            lblTitle.Text = "QS Toolbox (QSBar)";
            lblTitle.Font = new Font("Microsoft YaHei", 18, FontStyle.Bold);
            lblTitle.Location = new Point(20, 20);
            lblTitle.Size = new Size(600, 45);
            lblTitle.ForeColor = Color.FromArgb(0, 120, 215);

            lblVersion.Location = new Point(25, 75);
            lblVersion.Size = new Size(585, 30);

            lblUpdateDate.Location = new Point(25, 105);
            lblUpdateDate.Size = new Size(585, 30);

            txtContent.Multiline = true;
            txtContent.ReadOnly = true;
            txtContent.ScrollBars = ScrollBars.Vertical;
            txtContent.BorderStyle = BorderStyle.None;
            txtContent.BackColor = Color.White;
            txtContent.Location = new Point(25, 145);
            txtContent.Size = new Size(585, 565);
            txtContent.Text = WpsExcelAddIn.UseChineseRibbon ?
                "快捷键:\r\n" +
                "--------------------------------------------------\r\n" +
                "Ctrl + 3:  批量处理\r\n" +
                "Ctrl + 4:  选择非空单元格\r\n" +
                "Ctrl + 5:  选择可见单元格\r\n" +
                "Ctrl + 6:  会计格式（0位小数）\r\n" +
                "Ctrl + 7:  会计格式（2位小数）\r\n" +
                "Ctrl + 8:  会计格式（3位小数）\r\n" +
                "Ctrl + 9:  亿万元单位切换\r\n" +
                "Ctrl + 0:  手动/自动计算切换\r\n" +
                "Ctrl + Alt + 1:  展开数据透视表\r\n" +
                "Ctrl + Alt + 2:  折叠数据透视表\r\n\r\n" +
                "功能说明:\r\n" +
                "--------------------------------------------------\r\n" +
                "[批量处理] 大数据量表格的快速运算和转换。\r\n" +
                "[导出报表] 导出标准报表、内部报表等多种格式。\r\n" +
                "[格式工具] 一键数值化、公式锁定、换行等。\r\n" +
                "[分级工具] 自动根据内容显示分级及分级样式。\r\n" +
                "[图片工具] 批量缩放图片并全选所有图片。\r\n" +
                "[工作表/文件工具] 合并工作表并生成目录。\r\n" +
                "[计算模式] 在自动计算和手动计算之间切换。" :
                "Shortcuts:\r\n" +
                "--------------------------------------------------\r\n" +
                "Ctrl + 3:  Batch process\r\n" +
                "Ctrl + 4:  Select non-empty cells\r\n" +
                "Ctrl + 5:  Select visible cells\r\n" +
                "Ctrl + 6:  Accounting format (0 decimals)\r\n" +
                "Ctrl + 7:  Accounting format (2 decimals)\r\n" +
                "Ctrl + 8:  Accounting format (3 decimals)\r\n" +
                "Ctrl + 9:  Switch between billion / ten-thousand format\r\n" +
                "Ctrl + 0:  Toggle manual / automatic calculation\r\n" +
                "Ctrl + Alt + 1:  Expand pivot table\r\n" +
                "Ctrl + Alt + 2:  Collapse pivot table\r\n\r\n" +
                "Features:\r\n" +
                "--------------------------------------------------\r\n" +
                "[Batch Process] Fast calculation and conversion for large data sets.\r\n" +
                "[Export Reports] Export standard, internal, and other report formats.\r\n" +
                "[Format Tools] One-click values, formula lock, and line wrapping.\r\n" +
                "[Outline Tools] Automatic outline display and styling.\r\n" +
                "[Picture Tools] Resize pictures in bulk and select all pictures.\r\n" +
                "[Sheet / File Tools] Merge worksheets and generate indexes.\r\n" +
                "[Calculation Mode] Switch between automatic and manual calculation.";

            lnkDownload.Text = WpsExcelAddIn.UseChineseRibbon ? "下载: https://gitee.com/kevin137/qsbar-release" : "Download: https://gitee.com/kevin137/qsbar-release";
            lnkDownload.Location = new Point(25, 730);
            lnkDownload.Size = new Size(585, 30);
            lnkDownload.LinkClicked += (s, e) => {
                try { Process.Start("https://gitee.com/kevin137/qsbar-release"); }
                catch { MessageBox.Show(WpsExcelAddIn.UseChineseRibbon ? "无法打开链接。" : "Unable to open the link."); }
            };

            btnClose.Text = WpsExcelAddIn.UseChineseRibbon ? "关闭" : "Close";
            btnClose.Location = new Point(510, 770);
            btnClose.Size = new Size(100, 35);
            btnClose.Click += (s, e) => this.Close();

            this.Controls.AddRange(new Control[] { lblTitle, lblVersion, lblUpdateDate, txtContent, lnkDownload, btnClose });
        }

        private void LoadInfo()
        {
            var assembly = Assembly.GetExecutingAssembly();
            var version = assembly.GetName().Version;

            lblVersion.Text = (WpsExcelAddIn.UseChineseRibbon ? "版本: v" : "Version: v") + version.Major + "." + version.Minor + "." + version.Build + "." + version.Revision;

            try {
                var lastWriteTime = File.GetLastWriteTime(assembly.Location);
                lblUpdateDate.Text = (WpsExcelAddIn.UseChineseRibbon ? "编译时间: " : "Build time: ") + lastWriteTime.ToString("yyyy-MM-dd HH:mm:ss");
            }
            catch {
                lblUpdateDate.Text = WpsExcelAddIn.UseChineseRibbon ? "编译时间: 不可用" : "Build time: unavailable";
            }
        }
    }
}
