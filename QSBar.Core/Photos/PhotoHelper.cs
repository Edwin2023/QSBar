namespace QSBar.Core.Photos;

public static class PhotoHelper
{
    public static (double width, double height) Scale(double width, double height, double percent)
    {
        var p = percent / 100.0;
        return (width * p, height * p);
    }
}
