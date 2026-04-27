using System;
using System.Collections.Generic;
using Autodesk.AutoCAD.BoundaryRepresentation;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;
using AutoCAD_BoardSorter.Geometry;
using AutoCAD_BoardSorter.Models;
using BrepVertex = Autodesk.AutoCAD.BoundaryRepresentation.Vertex;

namespace AutoCAD_BoardSorter
{
    internal sealed class AssemblyPanelClassifier
    {
        private readonly BoardDimensionAnalyzer dimensionAnalyzer = new BoardDimensionAnalyzer();

        public bool TryClassify(
            Solid3d solid,
            AssemblyContainer container,
            SpecificationData specification,
            AssemblyPartData partData,
            out AssemblyPanel panel)
        {
            panel = null;
            if (solid == null || container == null)
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

            Point3d reference = container.Origin;
            List<Point3d> points = GetBrepVertices(solid);
            if (points.Count < 4)
            {
                Point3d min = extents.MinPoint;
                Point3d max = extents.MaxPoint;
                points.AddRange(new[]
                {
                    new Point3d(min.X, min.Y, min.Z),
                    new Point3d(min.X, min.Y, max.Z),
                    new Point3d(min.X, max.Y, min.Z),
                    new Point3d(min.X, max.Y, max.Z),
                    new Point3d(max.X, min.Y, min.Z),
                    new Point3d(max.X, min.Y, max.Z),
                    new Point3d(max.X, max.Y, min.Z),
                    new Point3d(max.X, max.Y, max.Z)
                });
            }

            double minX = double.MaxValue;
            double maxX = double.MinValue;
            double minY = double.MaxValue;
            double maxY = double.MinValue;
            double minZ = double.MaxValue;
            double maxZ = double.MinValue;

            foreach (Point3d corner in points)
            {
                Vector3d delta = corner - reference;
                double x = delta.DotProduct(container.FrontAxis);
                double y = delta.DotProduct(container.UpAxis);
                double z = delta.DotProduct(container.DepthAxis);
                minX = Math.Min(minX, x);
                maxX = Math.Max(maxX, x);
                minY = Math.Min(minY, y);
                maxY = Math.Max(maxY, y);
                minZ = Math.Min(minZ, z);
                maxZ = Math.Max(maxZ, z);
            }

            string role = NormalizeRole(partData == null ? null : partData.PartRole);
            double rawDepth = Math.Max(0.0, maxZ - minZ);

            minX = Clamp(minX, 0.0, container.Width);
            maxX = Clamp(maxX, 0.0, container.Width);
            minY = Clamp(minY, 0.0, container.Height);
            maxY = Clamp(maxY, 0.0, container.Height);
            minZ = Clamp(minZ, 0.0, container.Depth);
            maxZ = Clamp(maxZ, 0.0, container.Depth);

            double dx = Math.Max(0.0, maxX - minX);
            double dy = Math.Max(0.0, maxY - minY);
            double dz = Math.Max(0.0, maxZ - minZ);
            if (dz <= 0.0 && string.Equals(role, AssemblyConstants.DrawerFrontRole, StringComparison.OrdinalIgnoreCase) && rawDepth > 0.0)
            {
                minZ = 0.0;
                maxZ = Math.Min(container.Depth, Math.Max(1.0, rawDepth));
                dz = Math.Max(0.0, maxZ - minZ);
            }

            if (dx <= 0.0 || dy <= 0.0 || dz <= 0.0)
            {
                return false;
            }

            if (string.IsNullOrWhiteSpace(role))
            {
                role = ClassifyByDimensions(dx, dy, dz);
            }

            if (string.Equals(role, AssemblyConstants.UnknownPanelRole, StringComparison.OrdinalIgnoreCase))
            {
                return false;
            }

            double analyzedLength = Math.Max(dx, Math.Max(dy, dz));
            double analyzedWidth = dx + dy + dz - Math.Min(dx, Math.Min(dy, dz)) - Math.Max(dx, Math.Max(dy, dz));
            double analyzedThickness = Math.Min(dx, Math.Min(dy, dz));
            string method;
            double lengthMm;
            double widthMm;
            double thicknessMm;
            if (dimensionAnalyzer.TryGetDimensions(solid, out lengthMm, out widthMm, out thicknessMm, out method))
            {
                analyzedLength = lengthMm;
                analyzedWidth = widthMm;
                analyzedThickness = thicknessMm;
            }

            panel = new AssemblyPanel
            {
                ObjectId = solid.ObjectId,
                Handle = solid.Handle.ToString(),
                PartRole = role,
                Material = FirstNonBlank(partData == null ? null : partData.Material, specification == null ? null : specification.Material),
                Thickness = analyzedThickness,
                Length = analyzedLength,
                Width = analyzedWidth,
                MinX = minX,
                MaxX = maxX,
                MinY = minY,
                MaxY = maxY,
                MinDepth = minZ,
                MaxDepth = maxZ,
                IsRecognized = true
            };

            return true;
        }

        private static string ClassifyByDimensions(double dx, double dy, double dz)
        {
            double[] dims = { dx, dy, dz };
            Array.Sort(dims);
            if (dims[0] > dims[1] * 0.55)
            {
                return AssemblyConstants.UnknownPanelRole;
            }

            if (dx <= dy && dx <= dz)
            {
                return AssemblyConstants.VerticalPanelRole;
            }

            if (dy <= dx && dy <= dz)
            {
                return AssemblyConstants.HorizontalPanelRole;
            }

            return AssemblyConstants.BackPanelRole;
        }

        private static string NormalizeRole(string value)
        {
            string trimmed = (value ?? string.Empty).Trim();
            if (string.Equals(trimmed, AssemblyConstants.VerticalPanelRole, StringComparison.OrdinalIgnoreCase))
            {
                return AssemblyConstants.VerticalPanelRole;
            }

            if (string.Equals(trimmed, AssemblyConstants.HorizontalPanelRole, StringComparison.OrdinalIgnoreCase))
            {
                return AssemblyConstants.HorizontalPanelRole;
            }

            if (string.Equals(trimmed, AssemblyConstants.FrontPanelRole, StringComparison.OrdinalIgnoreCase))
            {
                return AssemblyConstants.FrontPanelRole;
            }

            if (string.Equals(trimmed, AssemblyConstants.BackPanelRole, StringComparison.OrdinalIgnoreCase))
            {
                return AssemblyConstants.BackPanelRole;
            }

            if (string.Equals(trimmed, AssemblyConstants.DrawerFrontRole, StringComparison.OrdinalIgnoreCase))
            {
                return AssemblyConstants.DrawerFrontRole;
            }

            if (string.Equals(trimmed, AssemblyConstants.DrawerSideRole, StringComparison.OrdinalIgnoreCase))
            {
                return AssemblyConstants.DrawerSideRole;
            }

            if (string.Equals(trimmed, AssemblyConstants.DrawerFrontWallRole, StringComparison.OrdinalIgnoreCase))
            {
                return AssemblyConstants.DrawerFrontWallRole;
            }

            if (string.Equals(trimmed, AssemblyConstants.DrawerBackWallRole, StringComparison.OrdinalIgnoreCase))
            {
                return AssemblyConstants.DrawerBackWallRole;
            }

            if (string.Equals(trimmed, AssemblyConstants.DrawerBottomRole, StringComparison.OrdinalIgnoreCase))
            {
                return AssemblyConstants.DrawerBottomRole;
            }

            return string.Empty;
        }

        private static double Clamp(double value, double min, double max)
        {
            return value < min ? min : (value > max ? max : value);
        }

        private static string FirstNonBlank(string first, string second)
        {
            if (!string.IsNullOrWhiteSpace(first))
            {
                return first.Trim();
            }

            return string.IsNullOrWhiteSpace(second) ? string.Empty : second.Trim();
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
    }
}
