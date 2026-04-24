using System;
using System.Globalization;
using Autodesk.AutoCAD.BoundaryRepresentation;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;
using BrepFace = Autodesk.AutoCAD.BoundaryRepresentation.Face;

namespace AutoCAD_BoardSorter.Geometry
{
    internal static class FaceKeyBuilder
    {
        private const double PositionPrecision = 0.001;
        private const double AreaPrecision = 0.001;

        public static bool TryBuild(FullSubentityPath path, out string key)
        {
            key = null;

            if (path.IsNull || path.SubentId.Type != SubentityType.Face)
            {
                return false;
            }

            try
            {
                using (var face = new BrepFace(path))
                {
                    key = Build(face);
                    return true;
                }
            }
            catch
            {
                return false;
            }
        }

        public static bool TryBuild(BrepFace face, out string key)
        {
            key = null;

            try
            {
                key = Build(face);
                return true;
            }
            catch
            {
                return false;
            }
        }

        public static bool IsFingerprintKey(string key)
        {
            return !string.IsNullOrWhiteSpace(key) && key.StartsWith("FP:", StringComparison.Ordinal);
        }

        public static bool TryGetBoundingDimensions(string key, out double first, out double second)
        {
            first = 0.0;
            second = 0.0;

            if (!IsFingerprintKey(key))
            {
                return false;
            }

            string[] parts = key.Split(':');
            if (parts.Length != 8)
            {
                return false;
            }

            double minX;
            double minY;
            double minZ;
            double maxX;
            double maxY;
            double maxZ;
            if (!double.TryParse(parts[2], NumberStyles.Float, CultureInfo.InvariantCulture, out minX)
                || !double.TryParse(parts[3], NumberStyles.Float, CultureInfo.InvariantCulture, out minY)
                || !double.TryParse(parts[4], NumberStyles.Float, CultureInfo.InvariantCulture, out minZ)
                || !double.TryParse(parts[5], NumberStyles.Float, CultureInfo.InvariantCulture, out maxX)
                || !double.TryParse(parts[6], NumberStyles.Float, CultureInfo.InvariantCulture, out maxY)
                || !double.TryParse(parts[7], NumberStyles.Float, CultureInfo.InvariantCulture, out maxZ))
            {
                return false;
            }

            double[] dimensions =
            {
                Math.Abs(maxX - minX),
                Math.Abs(maxY - minY),
                Math.Abs(maxZ - minZ)
            };

            Array.Sort(dimensions);
            Array.Reverse(dimensions);
            first = dimensions[0];
            second = dimensions[1];
            return first > 0.0 || second > 0.0;
        }

        private static string Build(BrepFace face)
        {
            BoundBlock3d bounds = face.BoundBlock;
            Point3d min = bounds.GetMinimumPoint();
            Point3d max = bounds.GetMaximumPoint();

            return "FP:"
                + Quantize(face.GetArea(), AreaPrecision) + ":"
                + Quantize(min.X, PositionPrecision) + ":"
                + Quantize(min.Y, PositionPrecision) + ":"
                + Quantize(min.Z, PositionPrecision) + ":"
                + Quantize(max.X, PositionPrecision) + ":"
                + Quantize(max.Y, PositionPrecision) + ":"
                + Quantize(max.Z, PositionPrecision);
        }

        private static string Quantize(double value, double precision)
        {
            double rounded = Math.Round(value / precision) * precision;
            return rounded.ToString("0.###", CultureInfo.InvariantCulture);
        }
    }
}
