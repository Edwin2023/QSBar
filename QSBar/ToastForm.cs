using System;
using System.Drawing;
using System.Windows.Forms;

namespace QSBar
{
    public class ToastForm : Form
    {
        private Timer timer;
        private int duration = 1500; // ms
        private int tickCount = 0;
        private int step = 30; // ms per tick

        public ToastForm(string message)
        {
            this.FormBorderStyle = FormBorderStyle.None;
            this.StartPosition = FormStartPosition.Manual;
            this.ShowInTaskbar = false;
            this.TopMost = true;
            this.BackColor = Color.FromArgb(50, 50, 50);
            this.Opacity = 0;
            this.Size = new Size(300, 60);

            Label lbl = new Label();
            lbl.Text = message;
            lbl.ForeColor = Color.White;
            lbl.Font = new Font("Microsoft YaHei UI", 12F, FontStyle.Regular);
            lbl.AutoSize = false;
            lbl.Dock = DockStyle.Fill;
            lbl.TextAlign = ContentAlignment.MiddleCenter;
            this.Controls.Add(lbl);

            // Center horizontally, slightly above middle vertically
            Rectangle workingArea = Screen.PrimaryScreen.WorkingArea;
            this.Location = new Point(
                workingArea.Width / 2 - this.Width / 2,
                workingArea.Height / 2 - this.Height / 2 - 150
            );
        }

        protected override void OnLoad(EventArgs e)
        {
            base.OnLoad(e);
            timer = new Timer();
            timer.Interval = step;
            timer.Tick += Timer_Tick;
            timer.Start();
        }

        private void Timer_Tick(object sender, EventArgs e)
        {
            tickCount += step;
            
            if (tickCount < 300)
            {
                // Fade in and slide up
                if (this.Opacity < 0.9) this.Opacity += 0.1;
                this.Top -= 2;
            }
            else if (tickCount > duration - 300)
            {
                // Fade out and slide up
                this.Opacity -= 0.1;
                this.Top -= 2;
            }

            if (tickCount >= duration)
            {
                timer.Stop();
                this.Close();
            }
        }

        private static ToastForm currentToast;

        public static void ShowToast(string message)
        {
            if (currentToast != null && !currentToast.IsDisposed)
            {
                currentToast.Close();
            }
            currentToast = new ToastForm(message);
            currentToast.Show();
        }
    }
}