using System;
using System.Collections.Generic;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace QSBar
{
    public partial class frmSelectSheets : Form
    {
        public List<string> SelectedSheetNames { get; private set; }

        public frmSelectSheets(Excel.Workbook workbook)
        {
            SelectedSheetNames = new List<string>();
            InitializeComponent();
            LoadSheets(workbook);
        }

        private void LoadSheets(Excel.Workbook workbook)
        {
            chkListSheets.Items.Clear();
            foreach (Excel.Worksheet sheet in workbook.Worksheets)
            {
                string sheetName = sheet.Name;

                if (sheetName.Equals("MergeSheet", StringComparison.OrdinalIgnoreCase) ||
                    sheetName.StartsWith("MergeSheet(", StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                string displayName = sheetName;
                if (sheet.Visible != Excel.XlSheetVisibility.xlSheetVisible)
                {
                    displayName += " (Hidden)";
                }

                chkListSheets.Items.Add(displayName, true);
            }
        }

        private void btnOK_Click(object sender, EventArgs e)
        {
            SelectedSheetNames.Clear();
            foreach (object item in chkListSheets.CheckedItems)
            {
                string displayName = item.ToString();
                string actualName = displayName;
                if (displayName.EndsWith(" (Hidden)"))
                {
                    actualName = displayName.Substring(0, displayName.Length - 9);
                }
                SelectedSheetNames.Add(actualName);
            }

            if (SelectedSheetNames.Count == 0)
            {
                MessageBox.Show("Please select at least one worksheet.");
                return;
            }

            this.DialogResult = DialogResult.OK;
            this.Close();
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.Cancel;
            this.Close();
        }

        private void btnSelectAll_Click(object sender, EventArgs e)
        {
            for (int i = 0; i < chkListSheets.Items.Count; i++)
            {
                chkListSheets.SetItemChecked(i, true);
            }
        }

        private void btnSelectNone_Click(object sender, EventArgs e)
        {
            for (int i = 0; i < chkListSheets.Items.Count; i++)
            {
                chkListSheets.SetItemChecked(i, false);
            }
        }
    }
}
