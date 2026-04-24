using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using Autodesk.AutoCAD.BoundaryRepresentation;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;
using AutoCAD_BoardSorter.Models;
using BrepEdge = Autodesk.AutoCAD.BoundaryRepresentation.Edge;
using BrepFace = Autodesk.AutoCAD.BoundaryRepresentation.Face;
using BrepVertex = Autodesk.AutoCAD.BoundaryRepresentation.Vertex;

namespace AutoCAD_BoardSorter.Geometry
{
    internal sealed class BoardCoatingAnalyzer
    {
        private const double DirectionTolerance = 0.88;

        private sealed class FaceRecord
        {
            public string Key;
            public readonly HashSet<string> Keys = new HashSet<string>(StringComparer.Ordinal);
            public readonly HashSet<string> EdgeKeys = new HashSet<string>(StringComparer.Ordinal);
            public Vector3d Normal;
            public double Area;
            public Point3d Center;
        }

        private sealed class BoardAxes
        {
            public Vector3d ThicknessAxis;
            public Vector3d LengthAxis;
            public Vector3d WidthAxis;
            public double Length;
            public double Width;
        }

        private sealed class EdgeRecord
        {
            public string Key;
            public Point3d Start;
            public Point3d End;
        }

        public BoardCoatingSlots Analyze(Solid3d solid, Transaction tr, SpecificationData specification)
        {
            return Analyze(solid, tr, specification, null);
        }

        public BoardCoatingSlots Analyze(Solid3d solid, Transaction tr, SpecificationData specification, BoardSortLogger log)
        {
            var slots = new BoardCoatingSlots();
            Dictionary<string, string> coatings = FaceCoatingStorage.ReadFaceCoatings(solid, tr);
            if (log != null)
            {
                log.Info("Coating analyze solid=" + solid.Handle + " records=" + coatings.Count.ToString(CultureInfo.InvariantCulture));
                foreach (KeyValuePair<string, string> coating in coatings)
                {
                    log.Info("  coating key=" + coating.Key + " value=\"" + coating.Value + "\"");
                }
            }

            if (coatings.Count == 0)
            {
                return slots;
            }

            List<Point3d> vertices = GetVertices(solid);
            List<FaceRecord> faces = GetPlanarFaces(solid);
            if (log != null)
            {
                log.Info("  brep vertices=" + vertices.Count.ToString(CultureInfo.InvariantCulture)
                    + " planarFaces=" + faces.Count.ToString(CultureInfo.InvariantCulture));
                foreach (FaceRecord face in faces)
                {
                    log.Info("  face key=" + face.Key + " aliases=" + string.Join(",", face.Keys));
                }
            }

            if (vertices.Count < 4 || faces.Count == 0)
            {
                if (log != null) log.Info("  skipped: no usable brep geometry");
                return slots;
            }

            BoardAxes axes;
            if (!TryBuildAxes(solid, vertices, faces, out axes))
            {
                if (log != null) log.Info("  skipped: axes not built");
                return slots;
            }

            if (log != null)
            {
                log.Info("  axes length=" + axes.Length.ToString("0.###", CultureInfo.InvariantCulture)
                    + " width=" + axes.Width.ToString("0.###", CultureInfo.InvariantCulture)
                    + " rotate=" + ((specification != null && specification.RotateLengthWidth) ? "yes" : "no"));
            }

            foreach (KeyValuePair<string, string> pair in coatings)
            {
                if (string.IsNullOrWhiteSpace(pair.Value))
                {
                    continue;
                }

                FaceRecord face = FindFace(faces, pair.Key);
                if (face == null)
                {
                    if (log != null) log.Info("  unresolved coating face key=" + pair.Key);
                    continue;
                }

                ApplySlot(slots, axes, face.Normal, pair.Value, specification != null && specification.RotateLengthWidth);
                if (log != null) log.Info("  applied key=" + pair.Key + " normal=" + FormatVector(face.Normal) + " value=\"" + pair.Value + "\"");
            }

            if (log != null)
            {
                log.Info("  result P1=\"" + slots.P1 + "\" P2=\"" + slots.P2
                    + "\" L1=\"" + slots.L1 + "\" L2=\"" + slots.L2
                    + "\" W1=\"" + slots.W1 + "\" W2=\"" + slots.W2 + "\"");
            }

            return slots;
        }

        public BoardSketch BuildSketch(Solid3d solid, Transaction tr, SpecificationData specification, BoardSortLogger log)
        {
            Dictionary<string, string> coatings = FaceCoatingStorage.ReadFaceCoatings(solid, tr);
            var sketch = new BoardSketch();
            List<Point3d> vertices = GetVertices(solid);
            List<FaceRecord> faces = GetPlanarFaces(solid);
            if (vertices.Count < 4 || faces.Count == 0)
            {
                return sketch;
            }

            BoardAxes axes;
            if (!TryBuildAxes(solid, vertices, faces, out axes))
            {
                return sketch;
            }

            FaceRecord plateFace = faces
                .Where(x => Math.Abs(VectorMath.Normalize(x.Normal).DotProduct(axes.ThicknessAxis)) >= DirectionTolerance)
                .OrderByDescending(x => x.Area)
                .FirstOrDefault();
            if (plateFace == null)
            {
                return sketch;
            }

            using (var brep = new Brep(solid))
            {
                foreach (BrepFace face in brep.Faces)
                {
                    string faceKey;
                    if (!FaceKeyBuilder.TryBuild(face, out faceKey) || !plateFace.Keys.Contains(faceKey))
                    {
                        continue;
                    }

                    var edgeRecords = new List<EdgeRecord>();
                    foreach (BoundaryLoop loop in face.Loops)
                    {
                        if (loop == null || loop.IsNull)
                        {
                            continue;
                        }

                        foreach (BrepEdge edge in loop.Edges)
                        {
                            string edgeKey;
                            if (!TryGetEdgeKey(edge, out edgeKey))
                            {
                                continue;
                            }

                            edgeRecords.Add(new EdgeRecord
                            {
                                Key = edgeKey,
                                Start = edge.Vertex1.Point,
                                End = edge.Vertex2.Point
                            });
                        }
                    }

                    BuildSketchFromEdges(sketch, edgeRecords, faces, plateFace, coatings, axes);
                    if (log != null)
                    {
                        log.Info("  sketch points=" + sketch.Points.Count.ToString(CultureInfo.InvariantCulture)
                            + " edges=" + sketch.Edges.Count.ToString(CultureInfo.InvariantCulture));
                    }

                    return sketch;
                }
            }

            return sketch;
        }

        private static string FormatVector(Vector3d vector)
        {
            return vector.X.ToString("0.###", CultureInfo.InvariantCulture)
                + "," + vector.Y.ToString("0.###", CultureInfo.InvariantCulture)
                + "," + vector.Z.ToString("0.###", CultureInfo.InvariantCulture);
        }

        private static void ApplySlot(BoardCoatingSlots slots, BoardAxes axes, Vector3d normal, string coating, bool rotate)
        {
            Vector3d n = VectorMath.Normalize(normal);
            double p = n.DotProduct(axes.ThicknessAxis);
            double l = n.DotProduct(axes.LengthAxis);
            double w = n.DotProduct(axes.WidthAxis);

            if (Math.Abs(p) >= DirectionTolerance)
            {
                if (p >= 0.0)
                {
                    slots.P1 = MergeCoating(slots.P1, coating);
                }
                else
                {
                    slots.P2 = MergeCoating(slots.P2, coating);
                }

                return;
            }

            if (Math.Abs(w) >= Math.Abs(l) && Math.Abs(w) >= DirectionTolerance)
            {
                if (!rotate)
                {
                    if (w >= 0.0) slots.L1 = MergeCoating(slots.L1, coating);
                    else slots.L2 = MergeCoating(slots.L2, coating);
                }
                else
                {
                    if (w >= 0.0) slots.W1 = MergeCoating(slots.W1, coating);
                    else slots.W2 = MergeCoating(slots.W2, coating);
                }

                return;
            }

            if (Math.Abs(l) >= DirectionTolerance)
            {
                if (!rotate)
                {
                    if (l >= 0.0) slots.W1 = MergeCoating(slots.W1, coating);
                    else slots.W2 = MergeCoating(slots.W2, coating);
                }
                else
                {
                    if (l >= 0.0) slots.L1 = MergeCoating(slots.L1, coating);
                    else slots.L2 = MergeCoating(slots.L2, coating);
                }
            }
        }

        private static string MergeCoating(string existing, string coating)
        {
            if (string.IsNullOrWhiteSpace(existing))
            {
                return coating ?? string.Empty;
            }

            if (string.Equals(existing, coating, StringComparison.CurrentCultureIgnoreCase))
            {
                return existing;
            }

            return existing + " / " + coating;
        }

        private static FaceRecord FindFace(IEnumerable<FaceRecord> faces, string key)
        {
            if (string.IsNullOrWhiteSpace(key))
            {
                return null;
            }

            foreach (FaceRecord face in faces)
            {
                if (face.Key == key || face.Keys.Contains(key))
                {
                    return face;
                }
            }

            return null;
        }

        private static bool TryBuildAxes(Solid3d solid, List<Point3d> vertices, List<FaceRecord> faces, out BoardAxes axes)
        {
            axes = null;
            FaceRecord largest = faces.OrderByDescending(x => x.Area).FirstOrDefault();
            if (largest == null || largest.Normal.Length <= VectorMath.Eps)
            {
                return false;
            }

            Vector3d thicknessAxis = VectorMath.Normalize(largest.Normal);
            List<Vector3d> directions = GetLinearEdgeDirections(solid, thicknessAxis);
            if (directions.Count == 0)
            {
                return false;
            }

            Vector3d lengthAxis = Vector3d.XAxis;
            Vector3d widthAxis = Vector3d.YAxis;
            double bestMetric = -1.0;

            foreach (Vector3d candidate in directions)
            {
                Vector3d x = VectorMath.Normalize(candidate);
                Vector3d y = thicknessAxis.CrossProduct(x);
                if (y.Length <= VectorMath.Eps)
                {
                    continue;
                }

                y = VectorMath.Normalize(y);
                double lx = SizeAlong(vertices, x);
                double ly = SizeAlong(vertices, y);
                double metric = lx * ly;
                if (metric > bestMetric)
                {
                    bestMetric = metric;
                    lengthAxis = lx >= ly ? x : y;
                    widthAxis = lx >= ly ? y : x;
                }
            }

            if (bestMetric <= 0.0)
            {
                return false;
            }

            axes = new BoardAxes
            {
                ThicknessAxis = thicknessAxis,
                LengthAxis = lengthAxis,
                WidthAxis = widthAxis,
                Length = SizeAlong(vertices, lengthAxis),
                Width = SizeAlong(vertices, widthAxis)
            };
            return true;
        }

        private static List<Vector3d> GetLinearEdgeDirections(Solid3d solid, Vector3d normal)
        {
            var directions = new List<Vector3d>();

            using (var brep = new Brep(solid))
            {
                foreach (BrepEdge edge in brep.Edges)
                {
                    Point3d first = edge.Vertex1.Point;
                    Point3d second = edge.Vertex2.Point;
                    Vector3d direction = second - first;
                    if (direction.Length <= VectorMath.Eps)
                    {
                        continue;
                    }

                    direction = direction - normal.MultiplyBy(direction.DotProduct(normal));
                    if (direction.Length <= VectorMath.Eps)
                    {
                        continue;
                    }

                    AddUniqueDirection(directions, VectorMath.Normalize(direction));
                }
            }

            return directions;
        }

        private static void AddUniqueDirection(List<Vector3d> directions, Vector3d direction)
        {
            foreach (Vector3d existing in directions)
            {
                if (Math.Abs(existing.DotProduct(direction)) > 0.985)
                {
                    return;
                }
            }

            directions.Add(direction);
        }

        private static double SizeAlong(IEnumerable<Point3d> vertices, Vector3d axis)
        {
            double min = double.MaxValue;
            double max = double.MinValue;

            foreach (Point3d point in vertices)
            {
                double value = point.X * axis.X + point.Y * axis.Y + point.Z * axis.Z;
                min = Math.Min(min, value);
                max = Math.Max(max, value);
            }

            return max - min;
        }

        private static List<Point3d> GetVertices(Solid3d solid)
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

        private static void AddUniquePoint(List<Point3d> points, Point3d point)
        {
            string key = PointKey(point);
            foreach (Point3d existing in points)
            {
                if (PointKey(existing) == key)
                {
                    return;
                }
            }

            points.Add(point);
        }

        private static List<FaceRecord> GetPlanarFaces(Solid3d solid)
        {
            var faces = new List<FaceRecord>();
            using (var brep = new Brep(solid))
            {
                int ordinal = 0;
                foreach (BrepFace face in brep.Faces)
                {
                    ordinal++;
                    Vector3d normal;
                    if (!TryGetFaceNormal(face, out normal))
                    {
                        continue;
                    }

                    string key;
                    if (!FaceKeyBuilder.TryBuild(face, out key))
                    {
                        continue;
                    }

                    var record = new FaceRecord
                    {
                        Key = key,
                        Normal = normal,
                        Area = Math.Abs(face.GetArea()),
                        Center = GetFaceCenter(face)
                    };
                    AddFaceEdgeKeys(face, record.EdgeKeys);
                    record.Keys.Add(key);
                    record.Keys.Add("1:" + ordinal.ToString(CultureInfo.InvariantCulture) + ":" + ordinal.ToString(CultureInfo.InvariantCulture));
                    record.Keys.Add("F:" + ordinal.ToString(CultureInfo.InvariantCulture) + ":" + ordinal.ToString(CultureInfo.InvariantCulture));

                    try
                    {
                        SubentityId subentityId = face.SubentityPath.SubentId;
                        long index = subentityId.IndexPtr.ToInt64();
                        record.Keys.Add("1:" + index.ToString(CultureInfo.InvariantCulture) + ":" + index.ToString(CultureInfo.InvariantCulture));
                        record.Keys.Add("F:" + index.ToString(CultureInfo.InvariantCulture) + ":" + index.ToString(CultureInfo.InvariantCulture));
                    }
                    catch
                    {
                    }

                    faces.Add(record);
                }
            }

            return faces;
        }

        private static void AddFaceEdgeKeys(BrepFace face, ISet<string> keys)
        {
            foreach (BoundaryLoop loop in face.Loops)
            {
                if (loop == null || loop.IsNull)
                {
                    continue;
                }

                foreach (BrepEdge edge in loop.Edges)
                {
                    string edgeKey;
                    if (TryGetEdgeKey(edge, out edgeKey))
                    {
                        keys.Add(edgeKey);
                    }
                }
            }
        }

        private static bool TryGetEdgeKey(BrepEdge edge, out string key)
        {
            key = null;

            try
            {
                string a = PointKey(edge.Vertex1.Point);
                string b = PointKey(edge.Vertex2.Point);
                key = string.CompareOrdinal(a, b) <= 0 ? a + "|" + b : b + "|" + a;
                return true;
            }
            catch
            {
                return false;
            }
        }

        private static void BuildSketchFromEdges(BoardSketch sketch,
                                                 IList<EdgeRecord> edgeRecords,
                                                 IList<FaceRecord> faces,
                                                 FaceRecord plateFace,
                                                 IDictionary<string, string> coatings,
                                                 BoardAxes axes)
        {
            var pointIndexes = new Dictionary<string, int>(StringComparer.Ordinal);
            foreach (EdgeRecord edge in edgeRecords)
            {
                int start = AddSketchPoint(sketch, pointIndexes, Project(edge.Start, axes));
                int end = AddSketchPoint(sketch, pointIndexes, Project(edge.End, axes));
                sketch.Edges.Add(new BoardSketchEdge
                {
                    StartIndex = start,
                    EndIndex = end,
                    Coating = FindCoatingForEdge(edge.Key, faces, plateFace, coatings)
                });
            }
        }

        private static BoardSketchPoint Project(Point3d point, BoardAxes axes)
        {
            return new BoardSketchPoint
            {
                X = point.X * axes.LengthAxis.X + point.Y * axes.LengthAxis.Y + point.Z * axes.LengthAxis.Z,
                Y = point.X * axes.WidthAxis.X + point.Y * axes.WidthAxis.Y + point.Z * axes.WidthAxis.Z
            };
        }

        private static int AddSketchPoint(BoardSketch sketch, IDictionary<string, int> pointIndexes, BoardSketchPoint point)
        {
            string key = Quantize(point.X) + "," + Quantize(point.Y);
            int index;
            if (pointIndexes.TryGetValue(key, out index))
            {
                return index;
            }

            index = sketch.Points.Count;
            sketch.Points.Add(point);
            pointIndexes.Add(key, index);
            return index;
        }

        private static string FindCoatingForEdge(string edgeKey, IEnumerable<FaceRecord> faces, FaceRecord plateFace, IDictionary<string, string> coatings)
        {
            foreach (FaceRecord face in faces)
            {
                if (ReferenceEquals(face, plateFace) || !face.EdgeKeys.Contains(edgeKey))
                {
                    continue;
                }

                foreach (string key in face.Keys)
                {
                    string coating;
                    if (coatings.TryGetValue(key, out coating) && !string.IsNullOrWhiteSpace(coating))
                    {
                        return coating;
                    }
                }
            }

            return string.Empty;
        }

        private static bool TryGetFaceNormal(BrepFace face, out Vector3d normal)
        {
            normal = Vector3d.ZAxis;

            try
            {
                PlanarEntity planar = face.Surface as PlanarEntity;
                if (planar != null)
                {
                    PlanarEquationCoefficients coefficients = planar.Coefficients;
                    object boxed = coefficients;
                    normal = new Vector3d(
                        GetDoubleMember(boxed, "A", "a"),
                        GetDoubleMember(boxed, "B", "b"),
                        GetDoubleMember(boxed, "C", "c"));

                    if (normal.Length > VectorMath.Eps)
                    {
                        return true;
                    }
                }
            }
            catch
            {
            }

            return TryGetFaceNormalFromBounds(face, out normal);
        }

        private static bool TryGetFaceNormalFromBounds(BrepFace face, out Vector3d normal)
        {
            normal = Vector3d.ZAxis;

            try
            {
                BoundBlock3d bounds = face.BoundBlock;
                Point3d min = bounds.GetMinimumPoint();
                Point3d max = bounds.GetMaximumPoint();

                double dx = Math.Abs(max.X - min.X);
                double dy = Math.Abs(max.Y - min.Y);
                double dz = Math.Abs(max.Z - min.Z);
                double maxSize = Math.Max(dx, Math.Max(dy, dz));
                double tol = Math.Max(0.001, maxSize * 0.000001);

                if (dx <= dy && dx <= dz && dx <= tol)
                {
                    normal = Vector3d.XAxis;
                    return true;
                }

                if (dy <= dx && dy <= dz && dy <= tol)
                {
                    normal = Vector3d.YAxis;
                    return true;
                }

                if (dz <= dx && dz <= dy && dz <= tol)
                {
                    normal = Vector3d.ZAxis;
                    return true;
                }
            }
            catch
            {
            }

            return false;
        }

        private static double GetDoubleMember(object instance, params string[] names)
        {
            Type type = instance.GetType();
            for (int i = 0; i < names.Length; i++)
            {
                System.Reflection.PropertyInfo property = type.GetProperty(names[i]);
                if (property != null)
                {
                    return Convert.ToDouble(property.GetValue(instance, null), CultureInfo.InvariantCulture);
                }

                System.Reflection.FieldInfo field = type.GetField(names[i]);
                if (field != null)
                {
                    return Convert.ToDouble(field.GetValue(instance), CultureInfo.InvariantCulture);
                }
            }

            return 0.0;
        }

        private static Point3d GetFaceCenter(BrepFace face)
        {
            BoundBlock3d bounds = face.BoundBlock;
            Point3d min = bounds.GetMinimumPoint();
            Point3d max = bounds.GetMaximumPoint();
            return new Point3d(
                (min.X + max.X) / 2.0,
                (min.Y + max.Y) / 2.0,
                (min.Z + max.Z) / 2.0);
        }

        private static string PointKey(Point3d point)
        {
            return Quantize(point.X) + "," + Quantize(point.Y) + "," + Quantize(point.Z);
        }

        private static string Quantize(double value)
        {
            return (Math.Round(value / 0.001) * 0.001).ToString("0.###", CultureInfo.InvariantCulture);
        }
    }
}
