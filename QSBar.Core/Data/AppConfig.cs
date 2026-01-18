namespace QSBar.Core.Data;

public static class AppConfig
{
    public static double DefaultTaxRate { get; set; } = 0.13;
    public static Dictionary<string, double> TaxRates { get; } = new Dictionary<string, double>
    {
        { "VAT", 0.13 },
        { "Service", 0.06 }
    };
    public static Dictionary<string, double> DiscountRates { get; } = new Dictionary<string, double>
    {
        { "VIP", 0.1 },
        { "Standard", 0.0 }
    };
}
