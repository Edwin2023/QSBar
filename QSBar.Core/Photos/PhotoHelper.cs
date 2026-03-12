using System;

namespace QSBar.Core.Photos
{
    public static class PhotoHelper
    {
        public static Tuple<double, double> Scale(double width, double height, double percent)
        {
            var p = percent / 100.0;
            return Tuple.Create(width * p, height * p);
        }
    }
}
