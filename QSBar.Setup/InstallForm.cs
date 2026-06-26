using System;
using System.Drawing;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace QSBar.Setup
{
    public class InstallForm : Form
    {
        private ProgressBar _progressBar;
        private Label _statusLabel;
        private Label _titleLabel;
        private bool _isUninstall;

        public InstallForm(bool isUninstall)
        {
            _isUninstall = isUninstall;
            InitializeComponent();
        }

        private void InitializeComponent()
        {
            this.Text = _isUninstall ? "QSBar Uninstaller" : "QSBar Installer";
            this.Size = new Size(450, 200);
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.StartPosition = FormStartPosition.CenterScreen;
            this.BackColor = Color.White;

            _titleLabel = new Label
            {
                Text = _isUninstall ? "Uninstalling QSBar..." : "Installing QSBar...",
                Font = new Font("Microsoft YaHei", 12, FontStyle.Bold),
                Location = new Point(20, 20),
                AutoSize = true
            };

            _statusLabel = new Label
            {
                Text = "Preparing...",
                Font = new Font("Microsoft YaHei", 9),
                Location = new Point(20, 60),
                Size = new Size(400, 20)
            };

            _progressBar = new ProgressBar
            {
                Location = new Point(20, 90),
                Size = new Size(400, 25),
                Style = ProgressBarStyle.Continuous,
                Minimum = 0,
                Maximum = 100
            };

            this.Controls.Add(_titleLabel);
            this.Controls.Add(_statusLabel);
            this.Controls.Add(_progressBar);
        }

        public void UpdateProgress(int value, string status)
        {
            if (this.InvokeRequired)
            {
                this.Invoke(new Action(() => UpdateProgress(value, status)));
                return;
            }

            if (value >= 0) _progressBar.Value = value;
            if (!string.IsNullOrEmpty(status)) _statusLabel.Text = status;
        }

        public async Task ExecuteAction(Func<Action<int, string>, Task> action)
        {
            try
            {
                await action((progress, status) => UpdateProgress(progress, status));
                UpdateProgress(100, _isUninstall ? "Uninstall complete." : "Install complete.");
                await Task.Delay(500);
                this.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(string.Format("The operation failed: {0}", ex.Message), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                this.Close();
            }
        }
    }
}
