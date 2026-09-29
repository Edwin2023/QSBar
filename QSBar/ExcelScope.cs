using System;
using Excel = Microsoft.Office.Interop.Excel;

namespace QSBar
{
    /// <summary>
    /// 批量操作期间关掉屏幕刷新、自动计算和事件，退出时恢复原状态。
    ///
    /// 自动计算是这里最要紧的一项。删名称、逐格写值这类动作每做一次都会触发一遍
    /// 全表重算，几千个名称就足够把 Excel 拖成假死 —— 净化名称的删除阶段漏掉这一项，
    /// 表现就是「一点就卡住」。
    ///
    /// 只在需要改的时候才写属性：用户本来就在手动计算模式时不去碰它，退出时也不改回
    /// 自动。屏幕刷新和事件同理，原来就是关着的（嵌套调用）就不管。
    /// </summary>
    internal sealed class ExcelScope : IDisposable
    {
        private readonly Excel.Application _app;

        private bool _restoreUpdating;
        private bool _restoreCalc;
        private bool _restoreEvents;
        private Excel.XlCalculation _oldCalc;

        private bool _disposed;

        private ExcelScope(Excel.Application app)
        {
            _app = app;
        }

        public static ExcelScope Begin(Excel.Application app)
        {
            var scope = new ExcelScope(app);
            if (app == null) return scope;

            try
            {
                if (app.ScreenUpdating)
                {
                    app.ScreenUpdating = false;
                    scope._restoreUpdating = true;
                }
            }
            catch { }

            // 没有打开工作簿时 Calculation 读写都会抛，且不同宿主（Excel / WPS）
            // 抛的时机还不一样，两边都得包住
            try
            {
                Excel.XlCalculation current = app.Calculation;
                if (current != Excel.XlCalculation.xlCalculationManual)
                {
                    scope._oldCalc = current;
                    app.Calculation = Excel.XlCalculation.xlCalculationManual;
                    scope._restoreCalc = true;
                }
            }
            catch { }

            try
            {
                if (app.EnableEvents)
                {
                    app.EnableEvents = false;
                    scope._restoreEvents = true;
                }
            }
            catch { }

            return scope;
        }

        public void Dispose()
        {
            if (_disposed) return;
            _disposed = true;
            if (_app == null) return;

            try { _app.StatusBar = false; } catch { }

            // 按设置的反序还原。恢复自动计算要放在最后一步之前：一还原 Excel 就会
            // 立刻重算整个工作簿，此时屏幕刷新还关着，用户看不到中间态
            if (_restoreEvents) { try { _app.EnableEvents = true; } catch { } }
            if (_restoreCalc) { try { _app.Calculation = _oldCalc; } catch { } }
            if (_restoreUpdating) { try { _app.ScreenUpdating = true; } catch { } }
        }
    }
}
