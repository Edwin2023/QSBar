using System;

namespace QSBar.Core.Functions;

public static class BQFunctions
{
    public static double PriceWithTax(double price, double taxRate)
    {
        return price * (1.0 + taxRate);
    }

    public static double Discount(double price, double rate)
    {
        return price * (1.0 - rate);
    }

    public static double CeilTo(double value, double step)
    {
        if (step <= 0) return value;
        var k = System.Math.Ceiling(value / step);
        return k * step;
    }

    public static double FloorTo(double value, double step)
    {
        if (step <= 0) return value;
        var k = System.Math.Floor(value / step);
        return k * step;
    }

    public static double SumRange(object[,] arr)
    {
        double sum = 0;
        int rMax = arr.GetLength(0);
        int cMax = arr.GetLength(1);
        for (int r = 0; r < rMax; r++)
        {
            for (int c = 0; c < cMax; c++)
            {
                var v = arr[r, c];
                if (v is double d) sum += d;
                else
                {
                    double parsed;
                    if (double.TryParse(v?.ToString() ?? "", out parsed)) sum += parsed;
                }
            }
        }
        return sum;
    }

    public static double Markup(double price, double markupRate)
    {
        return price * (1.0 + markupRate);
    }

    public static double NetOfTax(double grossPrice, double taxRate)
    {
        return grossPrice / (1.0 + taxRate);
    }

    public static double RoundTo(double value, int decimals)
    {
        return System.Math.Round(value, decimals, System.MidpointRounding.AwayFromZero);
    }

    public static double WeightedAverage(object[,] prices, object[,] quantities)
    {
        int rMax = System.Math.Min(prices.GetLength(0), quantities.GetLength(0));
        int cMax = System.Math.Min(prices.GetLength(1), quantities.GetLength(1));
        double sumPQ = 0;
        double sumQ = 0;
        for (int r = 0; r < rMax; r++)
        {
            for (int c = 0; c < cMax; c++)
            {
                double p = 0;
                double q = 0;
                var vp = prices[r, c];
                var vq = quantities[r, c];
                if (vp is double dp) p = dp; else double.TryParse(vp?.ToString() ?? "", out p);
                if (vq is double dq) q = dq; else double.TryParse(vq?.ToString() ?? "", out q);
                sumPQ += p * q;
                sumQ += q;
            }
        }
        if (sumQ == 0) return 0;
        return sumPQ / sumQ;
    }

    public static double PercentageChange(double oldValue, double newValue)
    {
        if (oldValue == 0) return 0;
        return (newValue - oldValue) / oldValue;
    }

    public static double GrossMargin(double salePrice, double cost)
    {
        if (salePrice == 0) return 0;
        return (salePrice - cost) / salePrice;
    }

    public static double PriceFromMargin(double cost, double marginRate)
    {
        double denom = 1.0 - marginRate;
        if (denom <= 0) return cost;
        return cost / denom;
    }

    public static double CostFromMargin(double salePrice, double marginRate)
    {
        return salePrice * (1.0 - marginRate);
    }

    public static double TieredDiscount(double price, object[,] thresholds, object[,] rates)
    {
        int tRows = thresholds.GetLength(0);
        int tCols = thresholds.GetLength(1);
        int rRows = rates.GetLength(0);
        int rCols = rates.GetLength(1);
        int count = System.Math.Min(tRows * tCols, rRows * rCols);
        var pairs = new System.Collections.Generic.List<(double th, double rate)>(count);
        for (int i = 0; i < tRows; i++)
        {
            for (int j = 0; j < tCols; j++)
            {
                int idx = i * tCols + j;
                if (idx >= count) break;
                double th = 0;
                var vt = thresholds[i, j];
                if (vt is double dt) th = dt; else double.TryParse(vt?.ToString() ?? "", out th);
                int ri = idx / rCols;
                int rj = idx % rCols;
                double rr = 0;
                var vr = rates[ri, rj];
                if (vr is double dr) rr = dr; else double.TryParse(vr?.ToString() ?? "", out rr);
                pairs.Add((th, rr));
            }
        }
        pairs.Sort((a, b) => a.th.CompareTo(b.th));
        double rateToUse = 0;
        foreach (var p in pairs)
        {
            if (price >= p.th) rateToUse = p.rate;
            else break;
        }
        return Discount(price, rateToUse);
    }

    public static object[,] AddTaxRange(object[,] arr, double taxRate)
    {
        int rMax = arr.GetLength(0);
        int cMax = arr.GetLength(1);
        var result = new object[rMax, cMax];
        for (int r = 0; r < rMax; r++)
        {
            for (int c = 0; c < cMax; c++)
            {
                var v = arr[r, c];
                double d = 0;
                if (v is double dv) d = dv;
                else double.TryParse(v?.ToString() ?? "", out d);
                result[r, c] = d * (1.0 + taxRate);
            }
        }
        return result;
    }
}
