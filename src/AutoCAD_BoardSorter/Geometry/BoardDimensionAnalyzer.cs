using System;
using System.Collections;
using System.Collections.Generic;
using System.Reflection;
using Autodesk.AutoCAD.BoundaryRepresentation;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;
using BrepEdge = Autodesk.AutoCAD.BoundaryRepresentation.Edge;
using BrepFace = Autodesk.AutoCAD.BoundaryRepresentation.Face;
using BrepVertex = Autodesk.AutoCAD.BoundaryRepresentation.Vertex;

namespace AutoCAD_BoardSorter.Geometry
{
    internal sealed class BoardDimensionAnalyzer
    {
        private sealed class FaceBasis
        {
            public Vector3d Ux;
            public Vector3d Uy;
            public Vector3d Uz;
            public BrepFace Face;
            public bool HasFaceEdgeAngles;
            public readonly List<double> CandidateAngles = new List<double>();
        }

        public bool TryGetDimensions(Solid3d solid, out double lengthMm, out double widthMm, out double thicknessMm, out string method)
        {
            lengthMm = 0.0;
            widthMm = 0.0;
            thicknessMm = 0.0;
            method = string.Empty;

            try
            {
                List<Point3d> vertices = GetBrepVertices(solid);
                if (vertices.Count < 4)
                {
                    return TryGetMassPrincipalDimensions(solid, out lengthMm, out widthMm, out thicknessMm, out method)
                        || TryGetExtentsDimensions(solid, out lengthMm, out widthMm, out thicknessMm, out method);
                }

                FaceBasis basis;
                if (TryFindLargestPlanarFaceBasis(solid, out basis))
                {
                    CollectCandidateAnglesFromLargestFaceEdges(basis);
                    if (!basis.HasFaceEdgeAngles)
                    {
                        CollectCandidateAnglesFromBrepEdges(solid, basis);
                        AddCoarseCandidateAngles(basis);
                    }

                    if (basis.CandidateAngles.Count == 0)
                    {
                        basis.CandidateAngles.Add(0.0);
                    }
                    else
                    {
                        AddUniqueAngle(basis.CandidateAngles, 0.0);
                    }

                    double bestAngle;
                    double bestLx;
                    double bestLy;
                    double bestLz;
                    double bestMetric;

                    EvaluateCandidateSet(vertices, basis, out bestAngle, out bestLx, out bestLy, out bestLz, out bestMetric);
                    if (!basis.HasFaceEdgeAngles)
                    {
                        RefineAngleBySearch(vertices, basis, ref bestAngle, ref bestLx, ref bestLy, ref bestLz, ref bestMetric);
                    }

                    lengthMm = bestLx;
                    widthMm = bestLy;
                    thicknessMm = bestLz;
                    VectorMath.Sort3Desc(ref lengthMm, ref widthMm, ref thicknessMm);
                    method = basis.HasFaceEdgeAngles
                        ? "BRep largest face edges"
                        : "BRep planar face + rotate";
                    return true;
                }

                if (TryGetLinearEdgeOrientedDimensions(vertices, solid, out lengthMm, out widthMm, out thicknessMm, out method))
                {
                    return true;
                }

                return TryGetMassPrincipalDimensions(solid, out lengthMm, out widthMm, out thicknessMm, out method)
                    || TryGetExtentsDimensions(solid, out lengthMm, out widthMm, out thicknessMm, out method);
            }
            catch
            {
                List<Point3d> vertices = TryGetBrepVerticesSafe(solid);
                return TryGetLinearEdgeOrientedDimensions(vertices, solid, out lengthMm, out widthMm, out thicknessMm, out method)
                    || TryGetMassPrincipalDimensions(solid, out lengthMm, out widthMm, out thicknessMm, out method)
                    || TryGetExtentsDimensions(solid, out lengthMm, out widthMm, out thicknessMm, out method);
            }
        }

        private static bool TryFindLargestPlanarFaceBasis(Solid3d solid, out FaceBasis basis)
        {
            basis = null;
            double bestArea = -1.0;

            using (var brep = new Brep(solid))
            {
                foreach (BrepFace face in brep.Faces)
                {
                    Vector3d normal;
                    if (!TryGetFaceNormal(face, out normal))
                    {
                        continue;
                    }

                    double area = Math.Abs(face.GetArea());
                    if (area <= bestArea)
                    {
                        continue;
                    }

                    var candidate = BuildBasis(normal);
                    candidate.Face = face;
                    basis = candidate;
                    bestArea = area;
                }
            }

            return basis != null;
        }

        private static FaceBasis BuildBasis(Vector3d normal)
        {
            Vector3d uz = VectorMath.Normalize(normal);
            Vector3d reference = Math.Abs(uz.X) < 0.9 ? Vector3d.XAxis : Vector3d.YAxis;
            Vector3d ux = VectorMath.Normalize(reference.CrossProduct(uz));
            Vector3d uy = VectorMath.Normalize(uz.CrossProduct(ux));

            return new FaceBasis
            {
                Ux = ux,
                Uy = uy,
                Uz = uz
            };
        }

        private static bool TryGetFaceNormal(BrepFace face, out Vector3d normal)
        {
            normal = Vector3d.ZAxis;

            try
            {
                PlanarEntity planar = face.Surface as PlanarEntity;
                if (planar == null)
                {
                    return false;
                }

                PlanarEquationCoefficients coefficients = planar.Coefficients;
                normal = GetNormalFromCoefficients(coefficients);
                if (normal.Length <= VectorMath.Eps)
                {
                    Vector3d u = new Vector3d(planar, new Vector2d(1.0, 0.0));
                    Vector3d v = new Vector3d(planar, new Vector2d(0.0, 1.0));
                    normal = u.CrossProduct(v);
                }

                return normal.Length > VectorMath.Eps;
            }
            catch
            {
            }

            return false;
        }

        private static Vector3d GetNormalFromCoefficients(PlanarEquationCoefficients coefficients)
        {
            object boxed = coefficients;
            return new Vector3d(
                GetDoubleMember(boxed, "A", "a"),
                GetDoubleMember(boxed, "B", "b"),
                GetDoubleMember(boxed, "C", "c"));
        }

        private static double GetDoubleMember(object instance, params string[] names)
        {
            Type type = instance.GetType();
            for (int i = 0; i < names.Length; i++)
            {
                PropertyInfo property = type.GetProperty(names[i], BindingFlags.Public | BindingFlags.Instance);
                if (property != null)
                {
                    return Convert.ToDouble(property.GetValue(instance, null));
                }

                FieldInfo field = type.GetField(names[i], BindingFlags.Public | BindingFlags.Instance);
                if (field != null)
                {
                    return Convert.ToDouble(field.GetValue(instance));
                }
            }

            return 0.0;
        }

        private static void CollectCandidateAnglesFromLargestFaceEdges(FaceBasis basis)
        {
            if (basis.Face == null)
            {
                return;
            }

            int before = basis.CandidateAngles.Count;
            CollectCandidateAnglesFromLargestFaceEdgesByReflection(basis);

            basis.HasFaceEdgeAngles = basis.CandidateAngles.Count > before;
        }

        private static void CollectCandidateAnglesFromLargestFaceEdgesByReflection(FaceBasis basis)
        {
            object loops = GetPropertyValue(basis.Face, "BoundaryLoops")
                ?? GetPropertyValue(basis.Face, "Loops");

            foreach (object loop in EnumerateObjects(loops))
            {
                object edges = GetPropertyValue(loop, "Edges");
                foreach (object edge in EnumerateObjects(edges))
                {
                    BrepEdge brepEdge = edge as BrepEdge;
                    if (brepEdge != null)
                    {
                        AddEdgeCandidateAngle(brepEdge, basis);
                    }
                }
            }
        }

        private static void CollectCandidateAnglesFromBrepEdges(Solid3d solid, FaceBasis basis)
        {
            try
            {
                using (var brep = new Brep(solid))
                {
                    foreach (BrepEdge edge in brep.Edges)
                    {
                        AddEdgeCandidateAngle(edge, basis);
                    }
                }
            }
            catch
            {
            }
        }

        private static void AddEdgeCandidateAngle(BrepEdge edge, FaceBasis basis)
        {
            Vector3d dir;
            if (!TryGetEdgeDirection(edge, out dir))
            {
                return;
            }

            Vector3d dirInPlane = VectorMath.ProjectToPlane(dir, basis.Uz);
            if (dirInPlane.Length <= VectorMath.Eps)
            {
                return;
            }

            dirInPlane = dirInPlane.GetNormal();
            double angle = Math.Atan2(dirInPlane.DotProduct(basis.Uy), dirInPlane.DotProduct(basis.Ux));
            AddUniqueAngle(basis.CandidateAngles, angle);
        }

        private static void AddCoarseCandidateAngles(FaceBasis basis)
        {
            for (double angleDeg = 0.0; angleDeg < 180.0; angleDeg += VectorMath.CoarseStepDeg)
            {
                AddUniqueAngle(basis.CandidateAngles, VectorMath.DegToRad(angleDeg));
            }
        }

        private static bool TryGetEdgeDirection(BrepEdge edge, out Vector3d direction)
        {
            direction = Vector3d.XAxis;

            try
            {
                LinearEntity3d linear = edge.Curve as LinearEntity3d;
                if (linear != null)
                {
                    direction = linear.Direction;
                    return direction.Length > VectorMath.Eps;
                }
            }
            catch
            {
            }

            try
            {
                Point3d start = edge.Vertex1.Point;
                Point3d end = edge.Vertex2.Point;
                direction = end - start;
                return direction.Length > VectorMath.Eps;
            }
            catch
            {
                return false;
            }
        }

        private static bool TryGetLinearEdgeOrientedDimensions(
            List<Point3d> vertices,
            Solid3d solid,
            out double lengthMm,
            out double widthMm,
            out double thicknessMm,
            out string method)
        {
            lengthMm = 0.0;
            widthMm = 0.0;
            thicknessMm = 0.0;
            method = string.Empty;

            if (vertices == null || vertices.Count < 4)
            {
                return false;
            }

            try
            {
                List<Vector3d> directions = GetUniqueLinearEdgeDirections(solid);
                if (directions.Count < 3)
                {
                    return false;
                }

                double bestMetric = double.MaxValue;
                double bestA = 0.0;
                double bestB = 0.0;
                double bestC = 0.0;

                for (int i = 0; i < directions.Count; i++)
                {
                    for (int j = i + 1; j < directions.Count; j++)
                    {
                        Vector3d ux = directions[i];
                        Vector3d uySeed = directions[j];
                        Vector3d uz = ux.CrossProduct(uySeed);
                        if (uz.Length <= 0.1)
                        {
                            continue;
                        }

                        uz = VectorMath.Normalize(uz);
                        Vector3d uy = VectorMath.Normalize(uz.CrossProduct(ux));

                        double a = GetSizeAlongAxis(vertices, ux);
                        double b = GetSizeAlongAxis(vertices, uy);
                        double c = GetSizeAlongAxis(vertices, uz);

                        double metric = a * b * c;
                        if (metric < bestMetric)
                        {
                            bestMetric = metric;
                            bestA = a;
                            bestB = b;
                            bestC = c;
                        }
                    }
                }

                if (bestMetric == double.MaxValue)
                {
                    return false;
                }

                lengthMm = bestA;
                widthMm = bestB;
                thicknessMm = bestC;
                VectorMath.Sort3Desc(ref lengthMm, ref widthMm, ref thicknessMm);
                method = "BRep linear edge axes fallback";
                return true;
            }
            catch
            {
                return false;
            }
        }

        private static List<Vector3d> GetUniqueLinearEdgeDirections(Solid3d solid)
        {
            var directions = new List<Vector3d>();

            using (var brep = new Brep(solid))
            {
                foreach (BrepEdge edge in brep.Edges)
                {
                    Vector3d direction;
                    if (!TryGetEdgeDirection(edge, out direction))
                    {
                        continue;
                    }

                    if (direction.Length <= VectorMath.Eps)
                    {
                        continue;
                    }

                    AddUniqueDirection(directions, direction);
                }
            }

            return directions;
        }

        private static void AddUniqueDirection(List<Vector3d> directions, Vector3d direction)
        {
            direction = CanonicalDirection(VectorMath.Normalize(direction));

            for (int i = 0; i < directions.Count; i++)
            {
                double dot = Math.Abs(directions[i].DotProduct(direction));
                dot = Math.Min(1.0, Math.Max(-1.0, dot));
                double angle = Math.Acos(dot);
                if (angle <= VectorMath.DegToRad(VectorMath.AngleTolDeg))
                {
                    return;
                }
            }

            directions.Add(direction);
        }

        private static Vector3d CanonicalDirection(Vector3d direction)
        {
            double ax = Math.Abs(direction.X);
            double ay = Math.Abs(direction.Y);
            double az = Math.Abs(direction.Z);
            double main = ax >= ay && ax >= az ? direction.X : ay >= az ? direction.Y : direction.Z;
            return main < 0.0 ? direction.Negate() : direction;
        }

        private static List<Point3d> GetBrepVertices(Solid3d solid)
        {
            var points = new List<Point3d>();

            using (var brep = new Brep(solid))
            {
                foreach (BrepVertex vertex in brep.Vertices)
                {
                    AddUniquePoint(points, vertex.Point);
                }
            }

            return points;
        }

        private static List<Point3d> TryGetBrepVerticesSafe(Solid3d solid)
        {
            try
            {
                return GetBrepVertices(solid);
            }
            catch
            {
                return new List<Point3d>();
            }
        }

        private static void AddUniquePoint(List<Point3d> points, Point3d point)
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

        private static void AddUniqueAngle(List<double> angles, double angle)
        {
            angle = VectorMath.NormalizeAnglePi(angle);
            for (int i = 0; i < angles.Count; i++)
            {
                if (VectorMath.AngleDistancePi(angles[i], angle) <= VectorMath.DegToRad(VectorMath.AngleTolDeg))
                {
                    return;
                }
            }

            angles.Add(angle);
        }

        private static void EvaluateCandidateSet(
            List<Point3d> vertices,
            FaceBasis basis,
            out double bestAngle,
            out double bestLx,
            out double bestLy,
            out double bestLz,
            out double bestMetric)
        {
            bestAngle = 0.0;
            bestLx = 0.0;
            bestLy = 0.0;
            bestLz = 0.0;
            bestMetric = double.MaxValue;

            foreach (double angle in basis.CandidateAngles)
            {
                double lx;
                double ly;
                double lz;
                double metric;
                EvaluateAngle(vertices, basis, angle, out lx, out ly, out lz, out metric);

                if (metric < bestMetric)
                {
                    bestMetric = metric;
                    bestAngle = angle;
                    bestLx = lx;
                    bestLy = ly;
                    bestLz = lz;
                }
            }
        }

        private static void RefineAngleBySearch(
            List<Point3d> vertices,
            FaceBasis basis,
            ref double bestAngle,
            ref double bestLx,
            ref double bestLy,
            ref double bestLz,
            ref double bestMetric)
        {
            double stepDeg = VectorMath.StartStepDeg;

            while (stepDeg >= VectorMath.MinStepDeg)
            {
                bool improved = true;
                while (improved)
                {
                    improved = false;
                    double[] angles =
                    {
                        VectorMath.NormalizeAnglePi(bestAngle - VectorMath.DegToRad(stepDeg)),
                        VectorMath.NormalizeAnglePi(bestAngle + VectorMath.DegToRad(stepDeg))
                    };

                    for (int i = 0; i < angles.Length; i++)
                    {
                        double lx;
                        double ly;
                        double lz;
                        double metric;
                        EvaluateAngle(vertices, basis, angles[i], out lx, out ly, out lz, out metric);

                        if (metric < bestMetric)
                        {
                            bestMetric = metric;
                            bestAngle = angles[i];
                            bestLx = lx;
                            bestLy = ly;
                            bestLz = lz;
                            improved = true;
                        }
                    }
                }

                stepDeg /= 2.0;
            }
        }

        private static void EvaluateAngle(
            List<Point3d> vertices,
            FaceBasis basis,
            double angle,
            out double lx,
            out double ly,
            out double lz,
            out double metric)
        {
            Vector3d ux = VectorMath.RotateInPlane(basis.Ux, basis.Uy, angle);
            Vector3d uy = VectorMath.RotateInPlanePerp(basis.Ux, basis.Uy, angle);

            lx = GetSizeAlongAxis(vertices, ux);
            ly = GetSizeAlongAxis(vertices, uy);
            lz = GetSizeAlongAxis(vertices, basis.Uz);
            metric = lx * ly;
        }

        private static double GetSizeAlongAxis(List<Point3d> vertices, Vector3d axis)
        {
            double min = double.MaxValue;
            double max = double.MinValue;

            foreach (Point3d point in vertices)
            {
                double projection = point.GetAsVector().DotProduct(axis);
                if (projection < min)
                {
                    min = projection;
                }

                if (projection > max)
                {
                    max = projection;
                }
            }

            return Math.Abs(max - min);
        }

        private static bool TryGetMassPrincipalDimensions(Solid3d solid, out double lengthMm, out double widthMm, out double thicknessMm, out string method)
        {
            lengthMm = 0.0;
            widthMm = 0.0;
            thicknessMm = 0.0;
            method = string.Empty;

            try
            {
                object massProperties = solid.MassProperties;
                Vector3d axis0;
                Vector3d axis1;
                Vector3d axis2;

                if (!TryGetPrincipalAxis(massProperties, 0, out axis0)
                    || !TryGetPrincipalAxis(massProperties, 1, out axis1)
                    || !TryGetPrincipalAxis(massProperties, 2, out axis2))
                {
                    return false;
                }

                axis0 = VectorMath.Normalize(axis0);
                axis1 = VectorMath.Normalize(axis1);
                axis2 = VectorMath.Normalize(axis2);

                if (axis0.CrossProduct(axis1).DotProduct(axis2) < 0.0)
                {
                    axis2 = axis2.Negate();
                }

                Matrix3d toPrincipalCoordinates = Matrix3d.AlignCoordinateSystem(
                    Point3d.Origin,
                    axis0,
                    axis1,
                    axis2,
                    Point3d.Origin,
                    Vector3d.XAxis,
                    Vector3d.YAxis,
                    Vector3d.ZAxis);

                using (Solid3d clone = (Solid3d)solid.Clone())
                {
                    clone.TransformBy(toPrincipalCoordinates);
                    Extents3d extents = clone.GeometricExtents;
                    lengthMm = extents.MaxPoint.X - extents.MinPoint.X;
                    widthMm = extents.MaxPoint.Y - extents.MinPoint.Y;
                    thicknessMm = extents.MaxPoint.Z - extents.MinPoint.Z;
                    VectorMath.Sort3Desc(ref lengthMm, ref widthMm, ref thicknessMm);
                    method = "MassProperties principal axes fallback";
                    return true;
                }
            }
            catch
            {
                return false;
            }
        }

        private static bool TryGetPrincipalAxis(object massProperties, int index, out Vector3d axis)
        {
            axis = Vector3d.XAxis;

            try
            {
                PropertyInfo principalAxes = massProperties.GetType().GetProperty("PrincipalAxes", BindingFlags.Public | BindingFlags.Instance);
                if (principalAxes != null)
                {
                    ParameterInfo[] indexParameters = principalAxes.GetIndexParameters();
                    object directValue = indexParameters.Length == 1
                        ? principalAxes.GetValue(massProperties, new object[] { index })
                        : principalAxes.GetValue(massProperties, null);

                    if (directValue is Vector3d directAxis)
                    {
                        axis = directAxis;
                        return axis.Length > VectorMath.Eps;
                    }

                    object indexedValue = TryReadIndexedValue(directValue, index);
                    if (indexedValue is Vector3d indexedAxis)
                    {
                        axis = indexedAxis;
                        return axis.Length > VectorMath.Eps;
                    }
                }

                object axes = GetPropertyValue(massProperties, "PrincipalAxes");
                object value = null;

                if (axes != null)
                {
                    value = TryReadIndexedValue(axes, index);
                }

                if (value == null)
                {
                    PropertyInfo item = massProperties.GetType().GetProperty("Item", new[] { typeof(int) });
                    if (item != null)
                    {
                        value = item.GetValue(massProperties, new object[] { index });
                    }
                }

                if (value is Vector3d)
                {
                    axis = (Vector3d)value;
                    return axis.Length > VectorMath.Eps;
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

        private static object GetPropertyValue(object instance, string propertyName)
        {
            if (instance == null)
            {
                return null;
            }

            PropertyInfo property = instance.GetType().GetProperty(propertyName, BindingFlags.Public | BindingFlags.Instance);
            return property == null ? null : property.GetValue(instance, null);
        }

        private static IEnumerable<object> EnumerateObjects(object value)
        {
            IEnumerable enumerable = value as IEnumerable;
            if (enumerable == null)
            {
                yield break;
            }

            foreach (object item in enumerable)
            {
                yield return item;
            }
        }

        private static bool TryGetExtentsDimensions(Solid3d solid, out double lengthMm, out double widthMm, out double thicknessMm, out string method)
        {
            lengthMm = 0.0;
            widthMm = 0.0;
            thicknessMm = 0.0;
            method = string.Empty;

            try
            {
                Extents3d extents = solid.GeometricExtents;
                lengthMm = extents.MaxPoint.X - extents.MinPoint.X;
                widthMm = extents.MaxPoint.Y - extents.MinPoint.Y;
                thicknessMm = extents.MaxPoint.Z - extents.MinPoint.Z;
                VectorMath.Sort3Desc(ref lengthMm, ref widthMm, ref thicknessMm);
                method = "GeometricExtents fallback";
                return true;
            }
            catch
            {
                return false;
            }
        }
    }
}
