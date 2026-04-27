using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;
using AutoCAD_BoardSorter.Models;

namespace AutoCAD_BoardSorter
{
    internal sealed class AssemblyDrawerGenerator
    {
        public int CreateDrawers(Database db, AssemblyContainer container, AssemblyDrawerLayout layout)
        {
            if (db == null || container == null || layout == null || layout.Segments.Count == 0)
            {
                PaletteDebugLogger.Info(db, "AssemblyDrawerGenerator skipped: missing db/container/layout");
                return 0;
            }

            int created = 0;
            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                BlockTable blockTable = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForWrite);
                BlockTableRecord modelSpace = (BlockTableRecord)tr.GetObject(SymbolUtilityServices.GetBlockModelSpaceId(db), OpenMode.ForWrite);

                foreach (IGrouping<int, AssemblyDrawerSegment> drawerGroup in layout.Segments.GroupBy(x => x.DrawerIndex).OrderBy(x => x.Key))
                {
                    if (CreateDrawerBlock(db, tr, blockTable, modelSpace, container, drawerGroup.Key, drawerGroup.ToList()))
                    {
                        created++;
                    }
                }

                tr.Commit();
            }

            PaletteDebugLogger.Info(db, "AssemblyDrawerGenerator blocksCreated=" + created);
            return created;
        }

        private static bool CreateDrawerBlock(
            Database db,
            Transaction tr,
            BlockTable blockTable,
            BlockTableRecord modelSpace,
            AssemblyContainer container,
            int drawerIndex,
            IList<AssemblyDrawerSegment> segments)
        {
            if (segments == null || segments.Count == 0)
            {
                return false;
            }

            double baseX = segments.Min(x => x.MinX);
            double baseY = segments.Min(x => x.MinY);
            double baseZ = segments.Min(x => x.MinZ);
            string blockName = MakeUniqueBlockName(blockTable, container, drawerIndex);

            var blockRecord = new BlockTableRecord
            {
                Name = blockName,
                Origin = Point3d.Origin
            };
            ObjectId blockId = blockTable.Add(blockRecord);
            tr.AddNewlyCreatedDBObject(blockRecord, true);

            foreach (AssemblyDrawerSegment segment in segments)
            {
                CreateSegmentSolid(db, tr, blockRecord, container, segment, baseX, baseY, baseZ);
            }

            Point3d insertWorld = container.Origin
                + container.FrontAxis.MultiplyBy(baseX)
                + container.UpAxis.MultiplyBy(baseY)
                + container.DepthAxis.MultiplyBy(baseZ);
            var blockRef = new BlockReference(Point3d.Origin, blockId);
            blockRef.SetDatabaseDefaults();
            modelSpace.AppendEntity(blockRef);
            tr.AddNewlyCreatedDBObject(blockRef, true);
            blockRef.TransformBy(BuildBlockLocalToWorld(insertWorld, container));

            string material = FirstNonBlank(segments.FirstOrDefault(x => string.Equals(x.Role, AssemblyConstants.DrawerFrontRole, StringComparison.OrdinalIgnoreCase))?.Material, segments[0].Material);
            SpecificationStorage.Write(blockRef, new SpecificationData
            {
                AssemblyNumber = container.AssemblyNumber,
                PartName = "Ящик " + (drawerIndex + 1).ToString(CultureInfo.InvariantCulture),
                PartType = "Блок",
                Material = material,
                Note = blockName
            }, tr);

            AssemblyPartStorage.Write(blockRef, new AssemblyPartData
            {
                AssemblyNumber = container.AssemblyNumber,
                PartRole = "DrawerGroup",
                SourceContainerHandle = container.Handle,
                GeneratedByConstructor = "1",
                Material = material
            }, tr);

            PaletteDebugLogger.Info(db, "AssemblyDrawerGenerator block=" + blockName + " refId=" + blockRef.ObjectId.Handle);
            return true;
        }

        private static Matrix3d BuildBlockLocalToWorld(Point3d insertWorld, AssemblyContainer container)
        {
            Vector3d upAxis = container.UpAxis.GetNormal();
            Vector3d depthAxis = container.DepthAxis.GetNormal();
            Vector3d frontAxis = depthAxis.CrossProduct(upAxis);
            if (frontAxis.Length <= 1e-8)
            {
                frontAxis = container.FrontAxis.GetNormal();
            }
            else
            {
                frontAxis = frontAxis.GetNormal();
                if (frontAxis.DotProduct(container.FrontAxis) < 0.0)
                {
                    frontAxis = -frontAxis;
                }
            }

            upAxis = frontAxis.CrossProduct(depthAxis).GetNormal();
            return Matrix3d.AlignCoordinateSystem(
                Point3d.Origin,
                Vector3d.XAxis,
                Vector3d.YAxis,
                Vector3d.ZAxis,
                insertWorld,
                frontAxis,
                upAxis,
                depthAxis);
        }

        private static void CreateSegmentSolid(
            Database db,
            Transaction tr,
            BlockTableRecord blockRecord,
            AssemblyContainer container,
            AssemblyDrawerSegment segment,
            double baseX,
            double baseY,
            double baseZ)
        {
            double width = Math.Max(1.0, segment.MaxX - segment.MinX);
            double height = Math.Max(1.0, segment.MaxY - segment.MinY);
            double depth = Math.Max(1.0, segment.MaxZ - segment.MinZ);

            var solid = new Solid3d();
            solid.SetDatabaseDefaults();
            solid.CreateBox(width, height, depth);
            Point3d center = new Point3d(
                segment.MinX - baseX + (width * 0.5),
                segment.MinY - baseY + (height * 0.5),
                segment.MinZ - baseZ + (depth * 0.5));
            solid.TransformBy(Matrix3d.Displacement(center - Point3d.Origin));

            if (segment.HasSideGroove)
            {
                TrySubtractGroove(solid, segment, baseX, baseY, baseZ);
            }

            blockRecord.AppendEntity(solid);
            tr.AddNewlyCreatedDBObject(solid, true);

            string material = (segment.Material ?? string.Empty).Trim();
            SpecificationStorage.Write(solid, new SpecificationData
            {
                AssemblyNumber = container.AssemblyNumber,
                PartName = segment.Name,
                PartType = "Площадной",
                LengthMm = Math.Max(width, Math.Max(height, depth)),
                WidthMm = width + height + depth - Math.Min(width, Math.Min(height, depth)) - Math.Max(width, Math.Max(height, depth)),
                RotateLengthWidth = false,
                Material = material,
                Note = string.Empty
            }, tr);

            AssemblyPartStorage.Write(solid, new AssemblyPartData
            {
                AssemblyNumber = container.AssemblyNumber,
                PartRole = segment.Role,
                SourceContainerHandle = container.Handle,
                GeneratedByConstructor = "1",
                Material = material
            }, tr);

            PaletteDebugLogger.Info(
                db,
                "AssemblyDrawerGenerator block segment role=" + segment.Role
                + " name=\"" + segment.Name + "\""
                + " size=" + width.ToString("0.###", CultureInfo.InvariantCulture)
                + "x" + height.ToString("0.###", CultureInfo.InvariantCulture)
                + "x" + depth.ToString("0.###", CultureInfo.InvariantCulture));
        }

        private static void TrySubtractGroove(Solid3d side, AssemblyDrawerSegment segment, double baseX, double baseY, double baseZ)
        {
            double sideWidth = Math.Max(1.0, segment.MaxX - segment.MinX);
            double removeWidth = Math.Max(0.1, sideWidth - Math.Max(0.1, segment.GrooveThickness));
            double grooveHeight = Math.Max(0.1, segment.GrooveMaxY - segment.GrooveMinY);
            double depth = Math.Max(1.0, segment.MaxZ - segment.MinZ);
            if (removeWidth <= 0.1 || grooveHeight <= 0.1)
            {
                return;
            }

            try
            {
                using (var cutter = new Solid3d())
                {
                    cutter.CreateBox(removeWidth, grooveHeight, depth + 0.1);
                    double cutterMinX = segment.GrooveOnMaxX
                        ? segment.MaxX - removeWidth
                        : segment.MinX;
                    Point3d cutterCenter = new Point3d(
                        cutterMinX - baseX + (removeWidth * 0.5),
                        segment.GrooveMinY - baseY + (grooveHeight * 0.5),
                        segment.MinZ - baseZ + (depth * 0.5));
                    cutter.TransformBy(Matrix3d.Displacement(cutterCenter - Point3d.Origin));
                    side.BooleanOperation(BooleanOperationType.BoolSubtract, cutter);
                }
            }
            catch
            {
            }
        }

        private static string MakeUniqueBlockName(BlockTable blockTable, AssemblyContainer container, int drawerIndex)
        {
            string prefix = "BD_DRAWER_" + SafeName(container.AssemblyNumber) + "_" + SafeName(container.Handle) + "_" + (drawerIndex + 1).ToString(CultureInfo.InvariantCulture);
            string name = prefix;
            int suffix = 1;
            while (blockTable.Has(name))
            {
                name = prefix + "_" + suffix.ToString(CultureInfo.InvariantCulture);
                suffix++;
            }

            return name;
        }

        private static string SafeName(string value)
        {
            string source = string.IsNullOrWhiteSpace(value) ? "A" : value.Trim();
            var chars = source.Select(ch => char.IsLetterOrDigit(ch) ? ch : '_').ToArray();
            return new string(chars);
        }

        private static string FirstNonBlank(string first, string second)
        {
            if (!string.IsNullOrWhiteSpace(first))
            {
                return first.Trim();
            }

            return string.IsNullOrWhiteSpace(second) ? string.Empty : second.Trim();
        }
    }
}
