using System;
using Autodesk.AutoCAD.Geometry;

namespace AutoCAD_BoardSorter.Geometry
{
    internal static class VectorMath
    {
        public const double Pi = Math.PI;
        public const double AngleTolDeg = 0.5;
        public const double MinStepDeg = 0.2;
        public const double StartStepDeg = 20.0;
        public const double CoarseStepDeg = 5.0;
        public const double Eps = 1e-8;

        public static Vector3d Normalize(Vector3d v)
        {
            if (v.Length <= Eps)
            {
                throw new InvalidOperationException("Zero vector.");
            }

            return v.GetNormal();
        }

        public static Vector3d ProjectToPlane(Vector3d v, Vector3d normal)
        {
            return v - normal.MultiplyBy(v.DotProduct(normal));
        }

        public static double DegToRad(double deg)
        {
            return deg * Pi / 180.0;
        }

        public static double NormalizeAnglePi(double angle)
        {
            while (angle < 0.0)
            {
                angle += Pi;
            }

            while (angle >= Pi)
            {
                angle -= Pi;
            }

            return angle;
        }

        public static double AngleDistancePi(double a, double b)
        {
            double d = Math.Abs(NormalizeAnglePi(a) - NormalizeAnglePi(b));
            if (d > Pi / 2.0)
            {
                d = Pi - d;
            }

            return Math.Abs(d);
        }

        public static void Sort3Desc(ref double a, ref double b, ref double c)
        {
            double t;
            if (a < b)
            {
                t = a;
                a = b;
                b = t;
            }

            if (b < c)
            {
                t = b;
                b = c;
                c = t;
            }

            if (a < b)
            {
                t = a;
                a = b;
                b = t;
            }
        }

        public static Vector3d RotateInPlane(Vector3d ux, Vector3d uy, double angle)
        {
            double c = Math.Cos(angle);
            double s = Math.Sin(angle);
            return Normalize(ux.MultiplyBy(c) + uy.MultiplyBy(s));
        }

        public static Vector3d RotateInPlanePerp(Vector3d ux, Vector3d uy, double angle)
        {
            double c = Math.Cos(angle);
            double s = Math.Sin(angle);
            return Normalize(ux.MultiplyBy(-s) + uy.MultiplyBy(c));
        }
    }
}
