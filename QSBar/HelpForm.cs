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

            this.Text = "About QS Toolbox";
            this.Size = new Size(550, 650);
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
            lblTitle.Size = new Size(500, 45);
            lblTitle.ForeColor = Color.FromArgb(0, 120, 215);

            lblVersion.Location = new Point(25, 75);
            lblVersion.Size = new Size(500, 30);

            lblUpdateDate.Location = new Point(25, 105);
            lblUpdateDate.Size = new Size(500, 30);

            txtContent.Multiline = true;
            txtContent.ReadOnly = true;
            txtContent.ScrollBars = ScrollBars.Vertical;
            txtContent.BorderStyle = BorderStyle.None;
            txtContent.BackColor = Color.White;
            txtContent.Location = new Point(25, 145);
            txtContent.Size = new Size(485, 380);
            txtContent.Text = "Shortcuts:\r\n" +
                              "--------------------------------------------------\r\n" +
                              "Ctrl + 3:  Batch process\r\n" +
                              "Ctrl + 4:  Select non-empty cells\r\n" +
                              "Ctrl + 5:  Select visible cells\r\n" +
                              "Ctrl + 6:  Accounting format (0 decimals)\r\n" +
                              "Ctrl + 7:  Accounting format (2 decimals)\r\n" +
                              "Ctrl + 8:  Accounting format (3 decimals)\r\n" +
                              "Ctrl + 9:  Switch between billion / ten-thousand format\r\n" +
                              "Ctrl + 0:  Toggle manual / automatic calculation\r\n\r\n" +
                              "Features:\r\n" +
                              "--------------------------------------------------\r\n" +
                              "[Batch Process] Fast calculation and conversion for large data sets.\r\n" +
                              "[Export Reports] Export standard, internal, and other report formats.\r\n" +
                              "[Format Tools] One-click values, formula lock, and line wrapping.\r\n" +
                              "[Outline Tools] Automatic outline display and styling.\r\n" +
                              "[Picture Tools] Resize pictures in bulk and select all pictures.\r\n" +
                              "[Sheet / File Tools] Merge worksheets and generate indexes.\r\n" +
                              "[Calculation Mode] Switch between automatic and manual calculation.";

            lnkDownload.Text = "Download: https://gitee.com/kevin137/qsbar-release";
            lnkDownload.Location = new Point(25, 540);
            lnkDownload.Size = new Size(500, 30);
            lnkDownload.LinkClicked += (s, e) => {
                try { Process.Start("https://gitee.com/kevin137/qsbar-release"); }
                catch { MessageBox.Show("Unable to open the link."); }
            };

            btnClose.Text = "Close";
            btnClose.Location = new Point(410, 565);
            btnClose.Size = new Size(100, 35);
            btnClose.Click += (s, e) => this.Close();

            this.Controls.AddRange(new Control[] { lblTitle, lblVersion, lblUpdateDate, txtContent, lnkDownload, btnClose });
        }

        private void LoadInfo()
        {
            var assembly = Assembly.GetExecutingAssembly();
            var version = assembly.GetName().Version;

            lblVersion.Text = string.Format("Version: v{0}.{1}.{2}.{3}", version.Major, version.Minor, version.Build, version.Revision);

            try {
                var lastWriteTime = File.GetLastWriteTime(assembly.Location);
                lblUpdateDate.Text = string.Format("Last build time: {0:yyyy-MM-dd HH:mm:ss}", lastWriteTime);
            }
            catch {
                lblUpdateDate.Text = "Build time: unavailable";
            }
        }
    }
}
