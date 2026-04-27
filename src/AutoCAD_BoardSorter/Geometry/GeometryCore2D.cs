using System;

namespace AutoCAD_BoardSorter.Geometry
{
    internal static class GeometryCore2D
    {
        public static double Clamp(double value, double min, double max)
        {
            return value < min ? min : (value > max ? max : value);
        }

        public static bool NearlyEqual(double first, double second, double tolerance)
        {
            return Math.Abs(first - second) <= tolerance;
        }
    }
}
