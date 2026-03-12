using System;
using System.Collections.Generic;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace QSBar
{
    public class BatchProcessForm : Form
    {

        private TextBox txtLeft;

        private TextBox txtCenter;

        private TextBox txtRight;

        private TextBox txtFind1;

        private TextBox txtRep1;

        private TextBox txtFind2;

        private TextBox txtRep2;

        private TextBox txtPreview;

        private Button btnOK;

        private Button btnCancel;



        public BatchProcessForm()

        {

            InitializeComponent();

            InitializeDefaults();

        }



        private void InitializeComponent()

        {

            this.Size = new Size(600, 750);

            this.Text = "批量处理";

            this.StartPosition = FormStartPosition.CenterScreen;

            this.FormBorderStyle = FormBorderStyle.Sizable;

            this.MaximizeBox = true;

            this.MinimizeBox = false;

            this.Font = new Font("Microsoft YaHei UI", 9F, FontStyle.Regular, GraphicsUnit.Point, ((byte)(134)));

            this.MinimumSize = new Size(600, 600);

            this.MaximumSize = new Size(600, 2000); // Lock width, allow height resizing



            int margin = 20;

            int y = 20;

            int labelWidth = 50;

            int textBoxWidth = 100;

            int rowHeight = 35;



            // Row 1: Left, Center, Right

            Label lblLeft = new Label { Text = "前缀:", Location = new Point(margin, y + 5), Width = labelWidth, AutoSize = false };

            txtLeft = new TextBox { Location = new Point(margin + labelWidth, y), Width = textBoxWidth };

            

            Label lblCenter = new Label { Text = "主体:", Location = new Point(margin + labelWidth + textBoxWidth + 20, y + 5), Width = labelWidth, AutoSize = false };

            txtCenter = new TextBox { Location = new Point(margin + labelWidth * 2 + textBoxWidth + 20, y), Width = textBoxWidth };

            

            Label lblRight = new Label { Text = "后缀:", Location = new Point(margin + (labelWidth + textBoxWidth + 20) * 2, y + 5), Width = labelWidth, AutoSize = false };

            txtRight = new TextBox { Location = new Point(margin + (labelWidth + textBoxWidth + 20) * 2 + labelWidth, y), Width = textBoxWidth };



            this.Controls.Add(lblLeft);

            this.Controls.Add(txtLeft);

            this.Controls.Add(lblCenter);

            this.Controls.Add(txtCenter);

            this.Controls.Add(lblRight);

            this.Controls.Add(txtRight);



            y += rowHeight + 10;



            // GroupBox for Find/Replace

            GroupBox grpReplace = new GroupBox { Text = "查找替换", Location = new Point(margin, y), Size = new Size(540, 120) };

            

            int grpY = 25;

            int findLabelWidth = 60;

            int findTextWidth = 170;

            int gap = 20;



            // Row 2: Replace 1 inside GroupBox

            Label lblFind1 = new Label { Text = "查找1:", Location = new Point(margin, grpY + 3), Width = findLabelWidth, AutoSize = false };

            txtFind1 = new TextBox { Location = new Point(margin + findLabelWidth, grpY), Width = findTextWidth };

            

            Label lblRep1 = new Label { Text = "替换1:", Location = new Point(margin + findLabelWidth + findTextWidth + gap, grpY + 3), Width = findLabelWidth, AutoSize = false };

            txtRep1 = new TextBox { Location = new Point(margin + findLabelWidth * 2 + findTextWidth + gap, grpY), Width = findTextWidth };



            grpReplace.Controls.Add(lblFind1);

            grpReplace.Controls.Add(txtFind1);

            grpReplace.Controls.Add(lblRep1);

            grpReplace.Controls.Add(txtRep1);



            grpY += rowHeight;



            // Row 3: Replace 2 inside GroupBox

            Label lblFind2 = new Label { Text = "查找2:", Location = new Point(margin, grpY + 3), Width = findLabelWidth, AutoSize = false };

            txtFind2 = new TextBox { Location = new Point(margin + findLabelWidth, grpY), Width = findTextWidth };



            Label lblRep2 = new Label { Text = "替换2:", Location = new Point(margin + findLabelWidth + findTextWidth + gap, grpY + 3), Width = findLabelWidth, AutoSize = false };

            txtRep2 = new TextBox { Location = new Point(margin + findLabelWidth * 2 + findTextWidth + gap, grpY), Width = findTextWidth };



            grpReplace.Controls.Add(lblFind2);

            grpReplace.Controls.Add(txtFind2);

            grpReplace.Controls.Add(lblRep2);

            grpReplace.Controls.Add(txtRep2);



            this.Controls.Add(grpReplace);



            y += 130;



            // Row 4: Preview Label

            Label lblPreview = new Label { Text = "预览结果:", Location = new Point(margin, y), Width = 200 };

            this.Controls.Add(lblPreview);



            y += 25;



            // Row 5: Preview TextBox

            // Dynamic height calculation

            int bottomMargin = 60;

            int previewHeight = this.ClientSize.Height - y - bottomMargin - margin;

            

            txtPreview = new TextBox

            {

                Location = new Point(margin, y),

                Size = new Size(540, previewHeight),

                Multiline = true,

                ScrollBars = ScrollBars.Vertical,

                ReadOnly = true,

                BackColor = Color.White,

                Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right

            };

            this.Controls.Add(txtPreview);



            // Row 6: Buttons

            int btnWidth = 100;

            int btnGap = 20;

            int btnY = this.ClientSize.Height - 45; // Fixed distance from bottom

            int btnX = (this.ClientSize.Width - (btnWidth * 2 + btnGap)) / 2;

            

            btnOK = new Button 

            { 

                Text = "确定", 

                Location = new Point(btnX, btnY), 

                Width = btnWidth, 

                Height = 35, 

                DialogResult = DialogResult.OK, 

                BackColor = SystemColors.ControlLight,

                Anchor = AnchorStyles.Bottom 

            };

            

            btnCancel = new Button 

            { 

                Text = "取消", 

                Location = new Point(btnX + btnWidth + btnGap, btnY), 

                Width = btnWidth, 

                Height = 35, 

                DialogResult = DialogResult.Cancel, 

                BackColor = SystemColors.ControlLight,

                Anchor = AnchorStyles.Bottom 

            };

            

            btnOK.Click += BtnOK_Click;

            

            this.Controls.Add(btnOK);

            this.Controls.Add(btnCancel);

            this.AcceptButton = btnOK;

            this.CancelButton = btnCancel;



            // Events

            EventHandler updateHandler = (s, e) => UpdatePreview();

            txtLeft.TextChanged += updateHandler;

            txtCenter.TextChanged += updateHandler;

            txtRight.TextChanged += updateHandler;

            txtFind1.TextChanged += updateHandler;

            txtRep1.TextChanged += updateHandler;

            txtFind2.TextChanged += updateHandler;

            txtRep2.TextChanged += updateHandler;

            

            // Activate event to refresh preview when form shows

            this.Shown += (s, e) => UpdatePreview();

        }



        private void InitializeDefaults()

        {

            txtCenter.Text = "<内容>";

        }



        private string SumUp(string everyValue, int item)

        {

            string lastValue;



            // Step 1: Center value logic

            if (txtCenter.Text.Contains("<内容>"))

            {

                lastValue = txtCenter.Text.Replace("<内容>", everyValue);

            }

            else

            {

                lastValue = txtCenter.Text;

            }



            // Step 2: Replace logic

            if (!string.IsNullOrEmpty(txtFind1.Text))

            {

                lastValue = lastValue.Replace(txtFind1.Text, txtRep1.Text);

            }

            if (!string.IsNullOrEmpty(txtFind2.Text))

            {

                lastValue = lastValue.Replace(txtFind2.Text, txtRep2.Text);

            }



            // Step 3: Concatenate

            lastValue = txtLeft.Text + lastValue + txtRight.Text;



            // Step 4: Index logic (#)

            if (item > 0)

            {

                lastValue = lastValue.Replace("#", item.ToString());

            }

            else

            {

                lastValue = lastValue.Replace("#", "");

            }



            return lastValue;

        }



        private void UpdatePreview()
        {
            try
            {
                Excel.Application app = WpsExcelAddIn.App;
                if (app == null) return;

                Excel.Range selection = app.Selection as Excel.Range;
                if (selection == null) return;

                StringBuilder sb = new StringBuilder();
                int item = 1;
                int count = 0;
                
                // Limit preview to first 50 items to avoid lag
                int maxPreview = 50; 

                // Use SpecialCells to get visible cells only
                try
                {
                    Excel.Range visibleCells = selection.SpecialCells(Excel.XlCellType.xlCellTypeVisible);
                    foreach (Excel.Range cell in visibleCells)
                    {
                        if (count >= maxPreview)
                        {
                            sb.AppendLine("...仅显示前 50 项");
                            break;
                        }

                        string formula = "";
                        try 
                        { 
                            object f = cell.Formula;
                            formula = (f != null) ? f.ToString() : ""; 
                        } 
                        catch { continue; } // Handle potential errors

                        string result = "";
                        if (formula.StartsWith("="))
                        {
                            result = "=" + SumUp(formula.Substring(1), item);
                        }
                        else
                        {
                            result = SumUp(formula, item);
                        }

                        sb.AppendLine(result);
                        
                        item++;
                        count++;
                    }
                }
                catch
                {
                    sb.AppendLine("无法获取选定区域的可见单元格");
                }

                txtPreview.Text = sb.ToString();
            }
            catch (Exception ex)
            {
                txtPreview.Text = "预览更新失败: " + ex.Message;
            }
        }



        private void BtnOK_Click(object sender, EventArgs e)
        {
            try
            {
                Excel.Application app = WpsExcelAddIn.App;
                if (app == null) return;

                Excel.Range selection = app.Selection as Excel.Range;
                if (selection == null) return;

                app.ScreenUpdating = false;
                try
                {
                    Excel.Range visibleCells = selection.SpecialCells(Excel.XlCellType.xlCellTypeVisible);
                    int item = 1;
                    var errorAddresses = new List<string>();
                    
                    foreach (Excel.Range cell in visibleCells)
                    {
                        string formula;
                        try 
                        { 
                            object f = cell.Formula;
                            formula = (f != null) ? f.ToString() : ""; 
                        }
                        catch
                        {
                            errorAddresses.Add(TryGetAddress(cell));
                            MarkCellRed(cell);
                            continue;
                        }

                        string result = "";
                        if (formula.StartsWith("="))
                        {
                            result = "=" + SumUp(formula.Substring(1), item);
                        }
                        else
                        {
                            result = SumUp(formula, item);
                        }

                        try
                        {
                            cell.Formula = result;
                            item++;
                        }
                        catch
                        {
                            errorAddresses.Add(TryGetAddress(cell));
                            MarkCellRed(cell);
                        }
                    }

                    if (errorAddresses.Count > 0)
                    {
                        MessageBox.Show(BuildErrorMessage(errorAddresses), "处理完成(部分错误)");
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("处理过程发生严重错误: " + ex.Message);
                }
                finally
                {
                    app.ScreenUpdating = true;
                }
                
                this.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show("错误: " + ex.Message);
            }
        }

        private static void MarkCellRed(Excel.Range cell)
        {
            try { cell.Font.Color = ColorTranslator.ToOle(Color.Red); } catch { }
        }

        private static string TryGetAddress(Excel.Range cell)
        {
            try { return (string)cell.Address[false, false]; } catch { }
            try { return (string)cell.Address; } catch { }
            return "<未知>";
        }



        private static string BuildErrorMessage(List<string> addresses)

        {

            const int max = 200;

            var sb = new StringBuilder();

            sb.AppendLine("以下单元格在写入新内容时发生错误（公式非法或单元格受保护）：");



            int take = Math.Min(max, addresses.Count);

            for (int i = 0; i < take; i++)

            {

                sb.Append(addresses[i]);

                if (i != take - 1) sb.Append(", ");

            }



            if (addresses.Count > max)

            {

                sb.AppendLine();

                sb.Append(string.Format("... 还有 {0} 个错误未列出", addresses.Count - max));

            }



            return sb.ToString();

        }

    }

}

