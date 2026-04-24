using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Reflection;
using System.Security.Cryptography;
using System.Text;
using Autodesk.AutoCAD.BoundaryRepresentation;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;
using BrepEdge = Autodesk.AutoCAD.BoundaryRepresentation.Edge;
using BrepVertex = Autodesk.AutoCAD.BoundaryRepresentation.Vertex;

namespace AutoCAD_BoardSorter.Geometry
{
    internal sealed class SolidFingerprintBuilder
    {
        private const double ToleranceMm = 0.05;

        public string Build(Solid3d solid, double lengthMm, double widthMm, double thicknessMm)
        {
            try
            {
                List<Point3d> vertices = GetVertices(solid);
                if (vertices.Count == 0)
                {
                    return BuildMassOnlyFingerprint(solid, lengthMm, widthMm, thicknessMm);
                }

                Basis basis = BuildPrincipalBasis(solid, vertices);
                var tokens = new List<string>();
                tokens.Add("D:" + Q(lengthMm) + "," + Q(widthMm) + "," + Q(thicknessMm));

                object mass = solid.MassProperties;
                tokens.Add("V:" + Q(GetDoubleProperty(mass, "Volume")));
                tokens.Add("A:" + Q(GetDoubleProperty(mass, "SurfaceArea", "Area")));

                using (var brep = new Brep(solid))
                {
                    tokens.Add("C:" + Count(brep.Faces) + "," + Count(brep.Edges) + "," + Count(brep.Vertices));

                    foreach (Point3d vertex in vertices)
                    {
                        tokens.Add("P:" + FormatPoint(ToLocal(vertex, basis)));
                    }

                    foreach (BrepEdge edge in brep.Edges)
                    {
                        tokens.Add(BuildEdgeToken(edge, basis));
                    }
                }

                tokens.Sort(StringComparer.Ordinal);
                return Hash(string.Join("|", tokens));
            }
            catch
            {
                return BuildMassOnlyFingerprint(solid, lengthMm, widthMm, thicknessMm);
            }
        }

        private static string BuildEdgeToken(BrepEdge edge, Basis basis)
        {
            try
            {
                Curve3d curve = edge.Curve;

                var circle = curve as CircularArc3d;
                if (circle != null)
                {
                    Point3d center = ToLocal(circle.Center, basis);
                    Vector3d normal = ToLocal(circle.Normal, basis);
                    return "CIR:" + FormatPoint(center) + ":R" + Q(circle.Radius) + ":N" + FormatVector(normal);
                }

                Point3d start = ToLocal(edge.Vertex1.Point, basis);
                Point3d end = ToLocal(edge.Vertex2.Point, basis);
                string a = FormatPoint(start);
                string b = FormatPoint(end);
                if (string.CompareOrdinal(a, b) > 0)
                {
                    string t = a;
                    a = b;
                    b = t;
                }

                return "E:" + a + ":" + b + ":L" + Q(start.DistanceTo(end));
            }
            catch
            {
                return "E:unknown";
            }
        }

        private static Basis BuildPrincipalBasis(Solid3d solid, List<Point3d> vertices)
        {
            Point3d origin = GetCentroid(solid, vertices);

            Vector3d x;
            Vector3d y;
            Vector3d z;
            if (!TryGetMassAxes(solid, out x, out y, out z))
            {
                Extents3d extents = solid.GeometricExtents;
                x = Vector3d.XAxis;
                y = Vector3d.YAxis;
                z = Vector3d.ZAxis;
                origin = new Point3d(
                    (extents.MinPoint.X + extents.MaxPoint.X) / 2.0,
                    (extents.MinPoint.Y + extents.MaxPoint.Y) / 2.0,
                    (extents.MinPoint.Z + extents.MaxPoint.Z) / 2.0);
            }

            x = CanonicalAxis(x);
            y = CanonicalAxis(y);
            z = CanonicalAxis(z);

            if (x.CrossProduct(y).DotProduct(z) < 0.0)
            {
                z = z.Negate();
            }

            return new Basis(origin, x, y, z);
        }

        private static bool TryGetMassAxes(Solid3d solid, out Vector3d x, out Vector3d y, out Vector3d z)
        {
            x = Vector3d.XAxis;
            y = Vector3d.YAxis;
            z = Vector3d.ZAxis;

            object mass = solid.MassProperties;
            return TryGetPrincipalAxis(mass, 0, out x)
                && TryGetPrincipalAxis(mass, 1, out y)
                && TryGetPrincipalAxis(mass, 2, out z)
                && x.Length > VectorMath.Eps
                && y.Length > VectorMath.Eps
                && z.Length > VectorMath.Eps;
        }

        private static Point3d GetCentroid(Solid3d solid, List<Point3d> vertices)
        {
            object mass = solid.MassProperties;
            object centroid = GetPropertyValue(mass, "Centroid");
            if (centroid is Point3d)
            {
                return (Point3d)centroid;
            }

            double x = 0.0;
            double y = 0.0;
            double z = 0.0;
            foreach (Point3d vertex in vertices)
            {
                x += vertex.X;
                y += vertex.Y;
                z += vertex.Z;
            }

            return new Point3d(x / vertices.Count, y / vertices.Count, z / vertices.Count);
        }

        private static Point3d ToLocal(Point3d point, Basis basis)
        {
            Vector3d v = point - basis.Origin;
            return new Point3d(v.DotProduct(basis.X), v.DotProduct(basis.Y), v.DotProduct(basis.Z));
        }

        private static Vector3d ToLocal(Vector3d vector, Basis basis)
        {
            return new Vector3d(vector.DotProduct(basis.X), vector.DotProduct(basis.Y), vector.DotProduct(basis.Z));
        }

        private static Vector3d CanonicalAxis(Vector3d axis)
        {
            axis = VectorMath.Normalize(axis);
            double ax = Math.Abs(axis.X);
            double ay = Math.Abs(axis.Y);
            double az = Math.Abs(axis.Z);

            double component = ax >= ay && ax >= az ? axis.X : ay >= az ? axis.Y : axis.Z;
            return component < 0.0 ? axis.Negate() : axis;
        }

        private static List<Point3d> GetVertices(Solid3d solid)
        {
            var points = new List<Point3d>();
            using (var brep = new Brep(solid))
            {
                foreach (BrepVertex vertex in brep.Vertices)
                {
                    AddUnique(points, vertex.Point);
                }
            }

            return points;
        }

        private static void AddUnique(List<Point3d> points, Point3d point)
        {
            for (int i = 0; i < points.Count; i++)
            {
                if (points[i].DistanceTo(point) <= 1e-6)
                {
                    return;
                }
            }

            points.Add(point);
        }

        private static int Count<T>(IEnumerable<T> items)
        {
            int count = 0;
            foreach (T item in items)
            {
                count++;
            }

            return count;
        }

        private static string BuildMassOnlyFingerprint(Solid3d solid, double lengthMm, double widthMm, double thicknessMm)
        {
            try
            {
                object mass = solid.MassProperties;
                return "M|" + Q(lengthMm) + "|" + Q(widthMm) + "|" + Q(thicknessMm)
                    + "|" + Q(GetDoubleProperty(mass, "Volume"))
                    + "|" + Q(GetDoubleProperty(mass, "SurfaceArea", "Area"));
            }
            catch
            {
                return "M|" + Q(lengthMm) + "|" + Q(widthMm) + "|" + Q(thicknessMm);
            }
        }

        private static bool TryGetPrincipalAxis(object massProperties, int index, out Vector3d axis)
        {
            axis = Vector3d.XAxis;

            try
            {
                object axes = GetPropertyValue(massProperties, "PrincipalAxes");
                object value = TryReadIndexedValue(axes, index);
                if (value is Vector3d)
                {
                    axis = (Vector3d)value;
                    return true;
                }
            }
            catch
            {
            }

            return false;
        }

        private static object TryReadIndexedValue(object instance, int index)
        {
            if (instance == null)
            {
                return null;
            }

            try
            {
                PropertyInfo item = instance.GetType().GetProperty("Item", new[] { typeof(int) });
                if (item != null)
                {
                    return item.GetValue(instance, new object[] { index });
                }

                if (instance is Array array && array.Length > index)
                {
                    return array.GetValue(index);
                }
            }
            catch
            {
            }

            return null;
        }

        private static double GetDoubleProperty(object instance, params string[] names)
        {
            for (int i = 0; i < names.Length; i++)
            {
                object value = GetPropertyValue(instance, names[i]);
                if (value != null)
                {
                    try
                    {
                        return Convert.ToDouble(value, CultureInfo.InvariantCulture);
                    }
                    catch
                    {
                    }
                }
            }

            return 0.0;
        }

        private static object GetPropertyValue(object instance, string propertyName)
        {
            if (instance == null)
            {
                return null;
            }

            PropertyInfo property = instance.GetType().GetProperty(propertyName, BindingFlags.Public | BindingFlags.Instance);
            return property == null ? null : property.GetValue(instance, null);
        }

        private static string FormatPoint(Point3d point)
        {
            return Q(point.X) + "," + Q(point.Y) + "," + Q(point.Z);
        }

        private static string FormatVector(Vector3d vector)
        {
            return Q(vector.X) + "," + Q(vector.Y) + "," + Q(vector.Z);
        }

        private static string Q(double value)
        {
            return Math.Round(value / ToleranceMm).ToString(CultureInfo.InvariantCulture);
        }

        private static string Hash(string value)
        {
            using (SHA256 sha = SHA256.Create())
            {
                byte[] bytes = sha.ComputeHash(Encoding.UTF8.GetBytes(value));
                return string.Concat(bytes.Select(x => x.ToString("x2", CultureInfo.InvariantCulture)));
            }
        }

        private readonly struct Basis
        {
            public readonly Point3d Origin;
            public readonly Vector3d X;
            public readonly Vector3d Y;
            public readonly Vector3d Z;

            public Basis(Point3d origin, Vector3d x, Vector3d y, Vector3d z)
            {
                Origin = origin;
                X = x;
                Y = y;
                Z = z;
            }
        }
    }
}
