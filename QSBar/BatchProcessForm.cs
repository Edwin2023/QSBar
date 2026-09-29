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

            this.Size = new Size(780, 800);

            this.Text = WpsExcelAddIn.UseChineseRibbon ? "批量处理" : "Batch Process";

            this.StartPosition = FormStartPosition.CenterScreen;

            this.FormBorderStyle = FormBorderStyle.Sizable;

            this.MaximizeBox = true;

            this.MinimizeBox = false;

            this.Font = new Font("Microsoft YaHei UI", 9F, FontStyle.Regular, GraphicsUnit.Point, ((byte)(134)));

            this.MinimumSize = new Size(780, 600);

            this.MaximumSize = new Size(780, 2000); // Lock width, allow height resizing



            int margin = 20;

            int y = 20;

            int gap = 15;

            int labelWidth = 65;

            int rowHeight = 35;

            // Three columns: Prefix (narrow) | Body (wide) | Suffix (narrow), evenly aligned to both edges

            int usableWidth = this.ClientSize.Width - margin * 2;

            int colNarrow = (usableWidth - gap * 2) * 32 / 100;  // ~32% each side

            int colWide = usableWidth - gap * 2 - colNarrow * 2;  // ~36% center

            int col1Start = margin;

            int col2Start = margin + colNarrow + gap;

            int col3Start = margin + colNarrow + gap + colWide + gap;



            // Row 1: Left, Center, Right

            Label lblLeft = new Label { Text = WpsExcelAddIn.UseChineseRibbon ? "前缀:" : "Prefix:", Location = new Point(col1Start, y + 5), Width = labelWidth, AutoSize = false };

            txtLeft = new TextBox { Location = new Point(col1Start + labelWidth, y), Width = colNarrow - labelWidth };



            Label lblCenter = new Label { Text = WpsExcelAddIn.UseChineseRibbon ? "主体:" : "Body:", Location = new Point(col2Start, y + 5), Width = labelWidth, AutoSize = false };

            txtCenter = new TextBox { Location = new Point(col2Start + labelWidth, y), Width = colWide - labelWidth };



            Label lblRight = new Label { Text = WpsExcelAddIn.UseChineseRibbon ? "后缀:" : "Suffix:", Location = new Point(col3Start, y + 5), Width = labelWidth, AutoSize = false };

            txtRight = new TextBox { Location = new Point(col3Start + labelWidth, y), Width = colNarrow - labelWidth };



            this.Controls.Add(lblLeft);

            this.Controls.Add(txtLeft);

            this.Controls.Add(lblCenter);

            this.Controls.Add(txtCenter);

            this.Controls.Add(lblRight);

            this.Controls.Add(txtRight);



            y += rowHeight + 10;



            // GroupBox for Find/Replace

            GroupBox grpReplace = new GroupBox { Text = WpsExcelAddIn.UseChineseRibbon ? "查找和替换" : "Find and Replace", Location = new Point(margin, y), Size = new Size(740, 120) };

            

            int grpY = 25;

            int findLabelWidth = 85;

            int findTextWidth = 255;



            // Row 2: Replace 1 inside GroupBox

            Label lblFind1 = new Label { Text = WpsExcelAddIn.UseChineseRibbon ? "查找1:" : "Find 1:", Location = new Point(margin, grpY + 3), Width = findLabelWidth, AutoSize = false };

            txtFind1 = new TextBox { Location = new Point(margin + findLabelWidth, grpY), Width = findTextWidth };

            

            Label lblRep1 = new Label { Text = WpsExcelAddIn.UseChineseRibbon ? "替换1:" : "Replace 1:", Location = new Point(margin + findLabelWidth + findTextWidth + gap, grpY + 3), Width = findLabelWidth, AutoSize = false };

            txtRep1 = new TextBox { Location = new Point(margin + findLabelWidth * 2 + findTextWidth + gap, grpY), Width = findTextWidth };



            grpReplace.Controls.Add(lblFind1);

            grpReplace.Controls.Add(txtFind1);

            grpReplace.Controls.Add(lblRep1);

            grpReplace.Controls.Add(txtRep1);



            grpY += rowHeight;



            // Row 3: Replace 2 inside GroupBox

            Label lblFind2 = new Label { Text = WpsExcelAddIn.UseChineseRibbon ? "查找2:" : "Find 2:", Location = new Point(margin, grpY + 3), Width = findLabelWidth, AutoSize = false };

            txtFind2 = new TextBox { Location = new Point(margin + findLabelWidth, grpY), Width = findTextWidth };



            Label lblRep2 = new Label { Text = WpsExcelAddIn.UseChineseRibbon ? "替换2:" : "Replace 2:", Location = new Point(margin + findLabelWidth + findTextWidth + gap, grpY + 3), Width = findLabelWidth, AutoSize = false };

            txtRep2 = new TextBox { Location = new Point(margin + findLabelWidth * 2 + findTextWidth + gap, grpY), Width = findTextWidth };



            grpReplace.Controls.Add(lblFind2);

            grpReplace.Controls.Add(txtFind2);

            grpReplace.Controls.Add(lblRep2);

            grpReplace.Controls.Add(txtRep2);



            this.Controls.Add(grpReplace);



            y += 130;



            // Row 4: Preview Label

            Label lblPreview = new Label { Text = WpsExcelAddIn.UseChineseRibbon ? "预览:" : "Preview:", Location = new Point(margin, y), Width = 200 };

            this.Controls.Add(lblPreview);



            y += 25;



            // Row 5: Preview TextBox

            // Dynamic height calculation

            int bottomMargin = 60;

            int previewHeight = this.ClientSize.Height - y - bottomMargin - margin;

            

            txtPreview = new TextBox

            {

                Location = new Point(margin, y),

                Size = new Size(740, previewHeight),

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

                Text = WpsExcelAddIn.UseChineseRibbon ? "确定" : "OK", 

                Location = new Point(btnX, btnY), 

                Width = btnWidth, 

                Height = 35, 

                DialogResult = DialogResult.OK, 

                BackColor = SystemColors.ControlLight,

                Anchor = AnchorStyles.Bottom 

            };

            

            btnCancel = new Button 

            { 

                Text = WpsExcelAddIn.UseChineseRibbon ? "取消" : "Cancel", 

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

            txtCenter.Text = WpsExcelAddIn.UseChineseRibbon ? "<内容>" : "<content>";

        }



        private string SumUp(string everyValue, int item)

        {

            string lastValue;



            // Step 1: Center value logic

            string contentTag = WpsExcelAddIn.UseChineseRibbon ? "<内容>" : "<content>";

            if (txtCenter.Text.Contains(contentTag))

            {

                lastValue = txtCenter.Text.Replace(contentTag, everyValue);

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
                Excel.Range visibleCells = null;
                try
                {
                    visibleCells = selection.SpecialCells(Excel.XlCellType.xlCellTypeVisible);
                    foreach (Excel.Range cell in visibleCells)
                    {
                        try
                        {
                            if (count >= maxPreview)
                            {
                                sb.AppendLine(WpsExcelAddIn.UseChineseRibbon ? "... 仅显示前50项" : "... only the first 50 items are shown");
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
                        finally { ComUtil.Release(cell); }
                    }
                }
                catch
                {
                    sb.AppendLine(WpsExcelAddIn.UseChineseRibbon ? "无法从选定范围获取可见单元格" : "Unable to get visible cells from the selected range");
                }
                finally { ComUtil.Release(visibleCells, selection); }

                txtPreview.Text = sb.ToString();
            }
            catch (Exception ex)
            {
                txtPreview.Text = (WpsExcelAddIn.UseChineseRibbon ? "预览更新失败: " : "Preview update failed: ") + ex.Message;
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

                // 逐格写公式，自动计算开着的话每写一格就重算一遍
                Excel.Range visibleCells = null;
                using (ExcelScope.Begin(app))
                try
                {
                    visibleCells = selection.SpecialCells(Excel.XlCellType.xlCellTypeVisible);
                    int item = 1;
                    var errorAddresses = new List<string>();

                    foreach (Excel.Range cell in visibleCells)
                    {
                        try
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
                        finally { ComUtil.Release(cell); }
                    }

                    if (errorAddresses.Count > 0)
                    {
                        MessageBox.Show(BuildErrorMessage(errorAddresses), WpsExcelAddIn.UseChineseRibbon ? "完成（部分错误）" : "Completed (Partial Errors)");
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show((WpsExcelAddIn.UseChineseRibbon ? "处理过程中发生致命错误: " : "A fatal error occurred during processing: ") + ex.Message);
                }
                finally
                {
                    ComUtil.Release(visibleCells, selection);
                }

                this.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show((WpsExcelAddIn.UseChineseRibbon ? "错误: " : "Error: ") + ex.Message);
            }
        }

        private static void MarkCellRed(Excel.Range cell)
        {
            Excel.Font font = null;
            try
            {
                font = cell.Font;
                font.Color = ColorTranslator.ToOle(Color.Red);
            }
            catch { }
            finally { ComUtil.Release(font); }
        }

        private static string TryGetAddress(Excel.Range cell)
        {
            try { return (string)cell.Address[false, false]; } catch { }
            try { return (string)cell.Address; } catch { }
            return "<unknown>";
        }



        private static string BuildErrorMessage(List<string> addresses)

        {

            const int max = 200;

            var sb = new StringBuilder();

            sb.AppendLine(WpsExcelAddIn.UseChineseRibbon ? "以下单元格写入新内容时失败（公式无效或单元格受保护）:" : "The following cells failed while writing new content (invalid formulas or protected cells):");



            int take = Math.Min(max, addresses.Count);

            for (int i = 0; i < take; i++)

            {

                sb.Append(addresses[i]);

                if (i != take - 1) sb.Append(", ");

            }



            if (addresses.Count > max)

            {

                sb.AppendLine();

                sb.Append(string.Format(WpsExcelAddIn.UseChineseRibbon ? "... 还有 {0} 个错误未列出" : "... {0} more errors not listed", addresses.Count - max));

            }



            return sb.ToString();

        }

    }

}

