using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using Autodesk.AutoCAD.BoundaryRepresentation;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;
using AutoCAD_BoardSorter.Models;
using BrepEdge = Autodesk.AutoCAD.BoundaryRepresentation.Edge;
using BrepVertex = Autodesk.AutoCAD.BoundaryRepresentation.Vertex;

namespace AutoCAD_BoardSorter
{
    internal sealed class AssemblyContainerAnalyzer
    {
        public bool TryAnalyze(Solid3d solid, AssemblyContainerData data, out AssemblyContainer container)
        {
            container = null;
            if (solid == null || data == null)
            {
                return false;
            }

            Vector3d frontAxis;
            Vector3d upAxis;
            Vector3d depthAxis;
            if (!TryReadAxes(data, out frontAxis, out upAxis, out depthAxis))
            {
                return false;
            }

            Extents3d extents;
            try
            {
                extents = solid.GeometricExtents;
            }
            catch
            {
                return false;
            }

            double storedWidth;
            double storedHeight;
            double storedDepth;
            if (TryParseDouble(data.Width, out storedWidth)
                && TryParseDouble(data.Height, out storedHeight)
                && TryParseDouble(data.Depth, out storedDepth)
                && TryAnalyzeRotatedBox(solid, data, frontAxis, upAxis, depthAxis, storedWidth, storedHeight, storedDepth, extents, out container))
            {
                return true;
            }

            if (TryAnalyzeRotatedBoxByStoredAxes(solid, data, frontAxis, upAxis, depthAxis, extents, out container))
            {
                return true;
            }

            Point3d reference = extents.MinPoint;
            List<Point3d> corners = GetBrepVertices(solid);
            if (corners.Count < 4)
            {
                corners = GetCorners(extents);
            }

            double minX;
            double maxX;
            double minY;
            double maxY;
            double minZ;
            double maxZ;
            GetRanges(corners, reference, frontAxis, upAxis, depthAxis, out minX, out maxX, out minY, out maxY, out minZ, out maxZ);

            Point3d origin = reference
                + frontAxis.MultiplyBy(minX)
                + upAxis.MultiplyBy(minY)
                + depthAxis.MultiplyBy(minZ);

            container = new AssemblyContainer
            {
                ObjectId = solid.ObjectId,
                Handle = solid.Handle.ToString(),
                AssemblyNumber = (data.AssemblyNumber ?? string.Empty).Trim(),
                FrontFaceKey = (data.FrontFaceKey ?? string.Empty).Trim(),
                FrontAxis = frontAxis,
                UpAxis = upAxis,
                DepthAxis = depthAxis,
                Bounds = extents,
                Origin = origin,
                Width = Math.Max(0.0, maxX - minX),
                Height = Math.Max(0.0, maxY - minY),
                Depth = Math.Max(0.0, maxZ - minZ)
            };

            return container.Width > 0.0 && container.Height > 0.0 && container.Depth > 0.0;
        }

        private static bool TryAnalyzeRotatedBoxByStoredAxes(
            Solid3d solid,
            AssemblyContainerData data,
            Vector3d storedFrontAxis,
            Vector3d storedUpAxis,
            Vector3d storedDepthAxis,
            Extents3d extents,
            out AssemblyContainer container)
        {
            container = null;
            List<Point3d> vertices = GetBrepVertices(solid);
            List<EdgeAxis> axes = GetBrepEdgeAxes(solid);
            if (vertices.Count < 4 || axes.Count < 3)
            {
                return false;
            }

            EdgeAxis frontEdge;
            EdgeAxis upEdge;
            EdgeAxis depthEdge;
            if (!TryPickAxisByDirection(axes, storedFrontAxis, null, out frontEdge)
                || !TryPickAxisByDirection(axes, storedUpAxis, new[] { frontEdge }, out upEdge)
                || !TryPickAxisByDirection(axes, storedDepthAxis, new[] { frontEdge, upEdge }, out depthEdge))
            {
                return false;
            }

            Vector3d frontAxis = OrientLike(frontEdge.Direction, storedFrontAxis);
            Vector3d upAxis = OrientLike(upEdge.Direction, storedUpAxis);
            Vector3d depthAxis = OrientLike(depthEdge.Direction, storedDepthAxis);
            return BuildContainerFromAxes(solid, data, extents, vertices, frontAxis, upAxis, depthAxis, out container);
        }

        private static bool TryAnalyzeRotatedBox(
            Solid3d solid,
            AssemblyContainerData data,
            Vector3d storedFrontAxis,
            Vector3d storedUpAxis,
            Vector3d storedDepthAxis,
            double storedWidth,
            double storedHeight,
            double storedDepth,
            Extents3d extents,
            out AssemblyContainer container)
        {
            container = null;
            List<Point3d> vertices = GetBrepVertices(solid);
            List<EdgeAxis> axes = GetBrepEdgeAxes(solid);
            if (vertices.Count < 4 || axes.Count < 3)
            {
                return false;
            }

            EdgeAxis frontEdge;
            EdgeAxis upEdge;
            EdgeAxis depthEdge;
            if (!TryPickAxis(axes, storedWidth, null, out frontEdge)
                || !TryPickAxis(axes, storedHeight, new[] { frontEdge }, out upEdge)
                || !TryPickAxis(axes, storedDepth, new[] { frontEdge, upEdge }, out depthEdge))
            {
                return false;
            }

            Vector3d upAxis = OrientLike(upEdge.Direction, storedUpAxis);
            Vector3d depthAxis = OrientLike(depthEdge.Direction, storedDepthAxis);
            Vector3d frontAxis = depthAxis.CrossProduct(upAxis);
            if (frontAxis.Length <= 1e-6)
            {
                return false;
            }

            frontAxis = frontAxis.GetNormal();
            if (Math.Abs(frontAxis.DotProduct(frontEdge.Direction)) < 0.75)
            {
                frontAxis = OrientLike(frontEdge.Direction, storedFrontAxis);
            }

            Point3d reference = vertices[0];
            return BuildContainerFromAxes(solid, data, extents, vertices, frontAxis, upAxis, depthAxis, out container);
        }

        private static bool BuildContainerFromAxes(
            Solid3d solid,
            AssemblyContainerData data,
            Extents3d extents,
            IList<Point3d> vertices,
            Vector3d frontAxis,
            Vector3d upAxis,
            Vector3d depthAxis,
            out AssemblyContainer container)
        {
            container = null;
            if (vertices == null || vertices.Count == 0)
            {
                return false;
            }

            Point3d reference = vertices[0];
            double minX;
            double maxX;
            double minY;
            double maxY;
            double minZ;
            double maxZ;
            GetRanges(vertices, reference, frontAxis, upAxis, depthAxis, out minX, out maxX, out minY, out maxY, out minZ, out maxZ);
            Point3d origin = reference
                + frontAxis.MultiplyBy(minX)
                + upAxis.MultiplyBy(minY)
                + depthAxis.MultiplyBy(minZ);

            container = new AssemblyContainer
            {
                ObjectId = solid.ObjectId,
                Handle = solid.Handle.ToString(),
                AssemblyNumber = (data.AssemblyNumber ?? string.Empty).Trim(),
                FrontFaceKey = (data.FrontFaceKey ?? string.Empty).Trim(),
                FrontAxis = frontAxis,
                UpAxis = upAxis,
                DepthAxis = depthAxis,
                Bounds = extents,
                Origin = origin,
                Width = Math.Max(0.0, maxX - minX),
                Height = Math.Max(0.0, maxY - minY),
                Depth = Math.Max(0.0, maxZ - minZ)
            };

            return container.Width > 0.0 && container.Height > 0.0 && container.Depth > 0.0;
        }

        public static string CreateFrontFaceKey(int axisIndex, bool useMax)
        {
            string axis = axisIndex == 0 ? "X" : (axisIndex == 1 ? "Y" : "Z");
            return axis + (useMax ? "Max" : "Min");
        }

        public static void BuildAxesFromFace(int axisIndex, bool useMax, out Vector3d frontAxis, out Vector3d upAxis, out Vector3d depthAxis)
        {
            Vector3d outward;
            switch (axisIndex)
            {
                case 0:
                    outward = useMax ? Vector3d.XAxis : -Vector3d.XAxis;
                    break;
                case 1:
                    outward = useMax ? Vector3d.YAxis : -Vector3d.YAxis;
                    break;
                default:
                    outward = useMax ? Vector3d.ZAxis : -Vector3d.ZAxis;
                    break;
            }

            depthAxis = -outward;
            Vector3d worldUp = Vector3d.ZAxis;
            Vector3d projectedUp = worldUp - depthAxis.MultiplyBy(worldUp.DotProduct(depthAxis));
            if (projectedUp.Length <= 1e-6)
            {
                worldUp = Vector3d.YAxis;
                projectedUp = worldUp - depthAxis.MultiplyBy(worldUp.DotProduct(depthAxis));
            }

            upAxis = projectedUp.GetNormal();
            frontAxis = depthAxis.CrossProduct(upAxis).GetNormal();
        }

        public static string FormatDouble(double value)
        {
            return value.ToString("0.###############", CultureInfo.InvariantCulture);
        }

        public static bool TryParseDouble(string value, out double result)
        {
            return double.TryParse(value ?? string.Empty, NumberStyles.Float, CultureInfo.InvariantCulture, out result);
        }

        private static bool TryReadAxes(AssemblyContainerData data, out Vector3d frontAxis, out Vector3d upAxis, out Vector3d depthAxis)
        {
            frontAxis = Vector3d.XAxis;
            upAxis = Vector3d.ZAxis;
            depthAxis = Vector3d.YAxis;

            double fx;
            double fy;
            double fz;
            double ux;
            double uy;
            double uz;
            double dx;
            double dy;
            double dz;
            if (!TryParseDouble(data.FrontAxisX, out fx)
                || !TryParseDouble(data.FrontAxisY, out fy)
                || !TryParseDouble(data.FrontAxisZ, out fz)
                || !TryParseDouble(data.UpAxisX, out ux)
                || !TryParseDouble(data.UpAxisY, out uy)
                || !TryParseDouble(data.UpAxisZ, out uz)
                || !TryParseDouble(data.DepthAxisX, out dx)
                || !TryParseDouble(data.DepthAxisY, out dy)
                || !TryParseDouble(data.DepthAxisZ, out dz))
            {
                return false;
            }

            frontAxis = new Vector3d(fx, fy, fz);
            upAxis = new Vector3d(ux, uy, uz);
            depthAxis = new Vector3d(dx, dy, dz);
            if (frontAxis.Length <= 1e-6 || upAxis.Length <= 1e-6 || depthAxis.Length <= 1e-6)
            {
                return false;
            }

            upAxis = upAxis.GetNormal();
            depthAxis = depthAxis.GetNormal();
            Vector3d resolvedFrontAxis = depthAxis.CrossProduct(upAxis);
            if (resolvedFrontAxis.Length <= 1e-6)
            {
                return false;
            }

            frontAxis = resolvedFrontAxis.GetNormal();
            return true;
        }

        private static List<Point3d> GetCorners(Extents3d extents)
        {
            Point3d min = extents.MinPoint;
            Point3d max = extents.MaxPoint;
            return new List<Point3d>
            {
                new Point3d(min.X, min.Y, min.Z),
                new Point3d(min.X, min.Y, max.Z),
                new Point3d(min.X, max.Y, min.Z),
                new Point3d(min.X, max.Y, max.Z),
                new Point3d(max.X, min.Y, min.Z),
                new Point3d(max.X, min.Y, max.Z),
                new Point3d(max.X, max.Y, min.Z),
                new Point3d(max.X, max.Y, max.Z)
            };
        }

        private static List<Point3d> GetBrepVertices(Solid3d solid)
        {
            var result = new List<Point3d>();
            try
            {
                using (var brep = new Brep(solid))
                {
                    foreach (BrepVertex vertex in brep.Vertices)
                    {
                        AddUniquePoint(result, vertex.Point);
                    }
                }
            }
            catch
            {
            }

            return result;
        }

        private static List<EdgeAxis> GetBrepEdgeAxes(Solid3d solid)
        {
            var result = new List<EdgeAxis>();
            try
            {
                using (var brep = new Brep(solid))
                {
                    foreach (BrepEdge edge in brep.Edges)
                    {
                        Point3d first = edge.Vertex1.Point;
                        Point3d second = edge.Vertex2.Point;
                        Vector3d vector = second - first;
                        if (vector.Length <= 1e-6)
                        {
                            continue;
                        }

                        AddEdgeAxis(result, vector.GetNormal(), vector.Length);
                    }
                }
            }
            catch
            {
            }

            return result.OrderByDescending(x => x.Length).ToList();
        }

        private static void AddUniquePoint(ICollection<Point3d> points, Point3d point)
        {
            foreach (Point3d existing in points)
            {
                if (existing.DistanceTo(point) <= 1e-4)
                {
                    return;
                }
            }

            points.Add(point);
        }

        private static void AddEdgeAxis(ICollection<EdgeAxis> axes, Vector3d direction, double length)
        {
            foreach (EdgeAxis axis in axes)
            {
                if (Math.Abs(axis.Direction.DotProduct(direction)) >= 0.96)
                {
                    axis.Length = Math.Max(axis.Length, length);
                    return;
                }
            }

            axes.Add(new EdgeAxis { Direction = direction, Length = length });
        }

        private static bool TryPickAxis(IEnumerable<EdgeAxis> axes, double targetLength, IEnumerable<EdgeAxis> used, out EdgeAxis axis)
        {
            axis = null;
            var usedList = used == null ? new List<EdgeAxis>() : used.ToList();
            double bestScore = double.MaxValue;
            foreach (EdgeAxis candidate in axes)
            {
                if (usedList.Any(x => Math.Abs(x.Direction.DotProduct(candidate.Direction)) >= 0.96))
                {
                    continue;
                }

                double score = Math.Abs(candidate.Length - targetLength);
                if (score < bestScore)
                {
                    bestScore = score;
                    axis = candidate;
                }
            }

            return axis != null;
        }

        private static bool TryPickAxisByDirection(IEnumerable<EdgeAxis> axes, Vector3d targetDirection, IEnumerable<EdgeAxis> used, out EdgeAxis axis)
        {
            axis = null;
            if (targetDirection.Length <= 1e-6)
            {
                return false;
            }

            Vector3d target = targetDirection.GetNormal();
            var usedList = used == null ? new List<EdgeAxis>() : used.ToList();
            double bestScore = double.MinValue;
            foreach (EdgeAxis candidate in axes)
            {
                if (usedList.Any(x => Math.Abs(x.Direction.DotProduct(candidate.Direction)) >= 0.96))
                {
                    continue;
                }

                double score = Math.Abs(candidate.Direction.DotProduct(target));
                if (score > bestScore)
                {
                    bestScore = score;
                    axis = candidate;
                }
            }

            return axis != null && bestScore >= 0.25;
        }

        private static Vector3d OrientLike(Vector3d direction, Vector3d preferred)
        {
            Vector3d normalized = direction.GetNormal();
            if (preferred.Length > 1e-6 && normalized.DotProduct(preferred.GetNormal()) < 0.0)
            {
                return -normalized;
            }

            return normalized;
        }

        private sealed class EdgeAxis
        {
            public Vector3d Direction;
            public double Length;
        }

        private static void GetRanges(
            IList<Point3d> corners,
            Point3d reference,
            Vector3d frontAxis,
            Vector3d upAxis,
            Vector3d depthAxis,
            out double minX,
            out double maxX,
            out double minY,
            out double maxY,
            out double minZ,
            out double maxZ)
        {
            minX = minY = minZ = double.MaxValue;
            maxX = maxY = maxZ = double.MinValue;

            foreach (Point3d corner in corners)
            {
                Vector3d delta = corner - reference;
                double x = delta.DotProduct(frontAxis);
                double y = delta.DotProduct(upAxis);
                double z = delta.DotProduct(depthAxis);
                minX = Math.Min(minX, x);
                maxX = Math.Max(maxX, x);
                minY = Math.Min(minY, y);
                maxY = Math.Max(maxY, y);
                minZ = Math.Min(minZ, z);
                maxZ = Math.Max(maxZ, z);
            }
        }
    }
}
