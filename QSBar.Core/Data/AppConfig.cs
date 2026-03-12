using System.Collections.Generic;

namespace QSBar.Core.Data
{
    public static class AppConfig
    {
        public static double DefaultTaxRate { get; set; }
        public static Dictionary<string, double> TaxRates { get; private set; }
        public static Dictionary<string, double> DiscountRates { get; private set; }

        static AppConfig()
        {
            DefaultTaxRate = 0.13;
            TaxRates = new Dictionary<string, double>
            {
                { "VAT", 0.13 },
                { "Service", 0.06 }
            };
            DiscountRates = new Dictionary<string, double>
            {
                { "VIP", 0.1 },
                { "Standard", 0.0 }
            };
        }
    }
}
