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
                if (input is double d)
                {
                    dt = DateTime.FromOADate(d);
                    success = true;
                }
                else if (input is DateTime date)
                {
                    dt = date;
                    success = true;
                }
                else
                {
                    string s = input.ToString();
                    if (string.IsNullOrWhiteSpace(s)) return "";
                    
                    if (DateTime.TryParse(s, out dt))
                    {
                        success = true;
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
