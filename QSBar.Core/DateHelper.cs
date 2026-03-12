using System;

namespace QSBar.Core
{
    public static class DateHelper
    {
        /// <summary>
        /// 将日期转换为 Excel 内部使用的 yyyy-M-d 格式 (如 2018-1-1)
        /// </summary>
        public static string FormatDate(object input)
        {
            if (input == null) return "";
            
            DateTime dt = DateTime.MinValue;
            bool success = false;
            
            try 
            {
                if (input is double)
                {
                    dt = DateTime.FromOADate((double)input);
                    success = true;
                }
                else if (input is DateTime)
                {
                    dt = (DateTime)input;
                    success = true;
                }
                else
                {
                    string s = input.ToString();
                    if (!string.IsNullOrWhiteSpace(s))
                    {
                        if (DateTime.TryParse(s, out dt))
                        {
                            success = true;
                        }
                    }
                }
            }
            catch
            {
                // Parse failed, success remains false
            }
            
            if (success)
            {
                return dt.ToString("yyyy-M-d");
            }
            
            // Fallback: simple replace if it looks like a date string with slashes
            return input.ToString().Replace("/", "-");
        }
    }
}
