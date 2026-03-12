using System;

namespace QSBar.Core.Functions
{
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
                    if (v is double)
                    {
                        sum += (double)v;
                    }
                    else
                    {
                        double parsed;
                        string s = (v != null) ? v.ToString() : "";
                        if (double.TryParse(s, out parsed)) sum += parsed;
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
                    
                    if (vp is double) p = (double)vp; 
                    else 
                    {
                        string sp = (vp != null) ? vp.ToString() : "";
                        double.TryParse(sp, out p);
                    }

                    if (vq is double) q = (double)vq; 
                    else 
                    {
                        string sq = (vq != null) ? vq.ToString() : "";
                        double.TryParse(sq, out q);
                    }

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
    }
}
