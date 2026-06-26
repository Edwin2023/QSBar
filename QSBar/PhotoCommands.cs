using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace QSBar
{
    public static class PhotoCommands
    {
        public static void ResizePictures(int times)
        {
            Excel.Application app = WpsExcelAddIn.App;
            if (app == null) return;
            
            try
            {
                // In Excel, selection can be a Range or a Drawing (ShapeRange)
                // We try to cast it to ShapeRange if possible, or use app.Selection
                
                object selection = app.Selection;
                if (selection == null) return;

                Excel.ShapeRange shapeRange = null;
                try
                {
                    // If multiple shapes are selected, selection is often the ShapeRange
                    shapeRange = selection as Excel.ShapeRange;
                    if (shapeRange == null)
                    {
                        // If one shape is selected, or we need to access it via property
                        // In some versions/cases, selection is a single object that has ShapeRange
                        dynamic dynSel = selection;
                        shapeRange = dynSel.ShapeRange;
                    }
                }
                catch { }

                if (shapeRange == null) return;

                // xlMoveAndSize = 1
                try { ((dynamic)selection).Placement = Excel.XlPlacement.xlMoveAndSize; } catch { }

                foreach (Excel.Shape obj in shapeRange)
                {
                    try
                    {
                        Excel.Range topLeft = obj.TopLeftCell;
                        
                        obj.Top = (float)topLeft.Top + 4;
                        obj.Left = (float)topLeft.Left + 4;
                        obj.LockAspectRatio = Microsoft.Office.Core.MsoTriState.msoFalse;
                        
                        obj.Height = (float)topLeft.Height * times - 8;
                        obj.Width = (float)topLeft.Width - 8;
                        
                        if (obj.Width < 10)
                        {
                            obj.Width = (float)topLeft.Width * 2 - 8;
                        }
                    }
                    catch { }
                }
            }
            catch { }
        }

        
        public static void SelectAllPictures()
        {
            Excel.Application app = WpsExcelAddIn.App;
            if (app == null) return;
            Excel.Worksheet activeSheet = app.ActiveSheet as Excel.Worksheet;
            
            if (activeSheet == null) return;

            if (activeSheet.ProtectContents)
            {
                MessageBox.Show("This worksheet is protected. The operation has been blocked.", "Notice");
                return;
            }

            try
            {
                activeSheet.Shapes.SelectAll();
            }
            catch { }
        }
    }
}

