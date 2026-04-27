using System;
using System.Windows;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;
using AutoCAD_BoardSorter.Models;

namespace AutoCAD_BoardSorter
{
    internal sealed class AssemblyPanelGenerator
    {
        public ObjectId CreatePanel(
            Database db,
            AssemblyContainer container,
            Rect modelRect,
            double availableDepth,
            double depthOffset,
            AssemblyEditorTool tool,
            double thickness,
            string material)
        {
            if (db == null || container == null || modelRect.IsEmpty)
            {
                PaletteDebugLogger.Info(db, "AssemblyPanelGenerator CreatePanel skipped: db/container/modelRect missing");
                return ObjectId.Null;
            }

            thickness = Math.Max(1.0, thickness);
            depthOffset = Math.Max(0.0, depthOffset);

            double minX = modelRect.Left;
            double maxX = modelRect.Right;
            double minY = modelRect.Top;
            double maxY = modelRect.Bottom;
            double minZ = 0.0;
            double maxDepth = Math.Max(1.0, Math.Min(container.Depth, availableDepth <= 0.0 ? container.Depth : availableDepth));
            double maxZ = maxDepth;
            string role = AssemblyConstants.UnknownPanelRole;
            string partName = "Деталь";

            switch (tool)
            {
                case AssemblyEditorTool.VerticalPanel:
                    role = AssemblyConstants.VerticalPanelRole;
                    partName = "Стойка";
                    break;
                case AssemblyEditorTool.HorizontalPanel:
                    role = AssemblyConstants.HorizontalPanelRole;
                    partName = "Полка";
                    break;
                case AssemblyEditorTool.FrontPanel:
                    minZ = depthOffset;
                    maxZ = Math.Min(maxDepth, depthOffset + thickness);
                    role = AssemblyConstants.FrontPanelRole;
                    partName = "Передняя стенка";
                    break;
                case AssemblyEditorTool.BackPanel:
                    maxZ = Math.Max(0.0, maxDepth - depthOffset);
                    minZ = Math.Max(0.0, maxZ - thickness);
                    role = AssemblyConstants.BackPanelRole;
                    partName = "Задняя стенка";
                    break;
                default:
                    PaletteDebugLogger.Info(db, "AssemblyPanelGenerator CreatePanel skipped: unsupported tool " + tool);
                    return ObjectId.Null;
            }

            double width = Math.Max(1.0, maxX - minX);
            double height = Math.Max(1.0, maxY - minY);
            double depth = Math.Max(1.0, maxZ - minZ);
            PaletteDebugLogger.Info(
                db,
                "AssemblyPanelGenerator CreatePanel role=" + role
                + " localBox=[" + minX.ToString("0.###", System.Globalization.CultureInfo.InvariantCulture)
                + "," + minY.ToString("0.###", System.Globalization.CultureInfo.InvariantCulture)
                + "," + minZ.ToString("0.###", System.Globalization.CultureInfo.InvariantCulture)
                + "]..[" + maxX.ToString("0.###", System.Globalization.CultureInfo.InvariantCulture)
                + "," + maxY.ToString("0.###", System.Globalization.CultureInfo.InvariantCulture)
                + "," + maxZ.ToString("0.###", System.Globalization.CultureInfo.InvariantCulture)
                + "] size=" + width.ToString("0.###", System.Globalization.CultureInfo.InvariantCulture)
                + "x" + height.ToString("0.###", System.Globalization.CultureInfo.InvariantCulture)
                + "x" + depth.ToString("0.###", System.Globalization.CultureInfo.InvariantCulture));
            ObjectId createdId = ObjectId.Null;

            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                BlockTableRecord modelSpace = (BlockTableRecord)tr.GetObject(SymbolUtilityServices.GetBlockModelSpaceId(db), OpenMode.ForWrite);
                var solid = new Solid3d();
                solid.SetDatabaseDefaults();
                solid.CreateBox(width, height, depth);
                modelSpace.AppendEntity(solid);
                tr.AddNewlyCreatedDBObject(solid, true);

                Point3d localCenter = container.Origin
                    + container.FrontAxis.MultiplyBy(minX + (width * 0.5))
                    + container.UpAxis.MultiplyBy(minY + (height * 0.5))
                    + container.DepthAxis.MultiplyBy(minZ + (depth * 0.5));
                Matrix3d localToWorld = Matrix3d.AlignCoordinateSystem(
                    Point3d.Origin,
                    Vector3d.XAxis,
                    Vector3d.YAxis,
                    Vector3d.ZAxis,
                    localCenter,
                    container.FrontAxis,
                    container.UpAxis,
                    container.DepthAxis);
                solid.TransformBy(localToWorld);

                SpecificationData spec;
                SpecificationStorage.TryRead(solid, tr, out spec);
                SpecificationStorage.Write(solid, new SpecificationData
                {
                    AssemblyNumber = container.AssemblyNumber,
                    PartName = partName,
                    PartType = "Площадной",
                    LengthMm = Math.Max(width, Math.Max(height, depth)),
                    WidthMm = width + height + depth - Math.Min(width, Math.Min(height, depth)) - Math.Max(width, Math.Max(height, depth)),
                    RotateLengthWidth = false,
                    Material = (material ?? string.Empty).Trim(),
                    Note = string.Empty
                }, tr);

                AssemblyPartStorage.Write(solid, new AssemblyPartData
                {
                    AssemblyNumber = container.AssemblyNumber,
                    PartRole = role,
                    SourceContainerHandle = container.Handle,
                    GeneratedByConstructor = "1",
                    Material = (material ?? string.Empty).Trim()
                }, tr);

                createdId = solid.ObjectId;
                tr.Commit();
            }

            PaletteDebugLogger.Info(db, "AssemblyPanelGenerator CreatePanel createdId=" + createdId.Handle);

            return createdId;
        }
    }
}
