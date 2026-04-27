using System;
using System.Collections.Generic;
using Autodesk.AutoCAD.BoundaryRepresentation;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;
using AutoCAD_BoardSorter.Models;
using BrepVertex = Autodesk.AutoCAD.BoundaryRepresentation.Vertex;

namespace AutoCAD_BoardSorter
{
    internal sealed class AssemblySceneLoader
    {
        private readonly AssemblyContainerAnalyzer containerAnalyzer = new AssemblyContainerAnalyzer();
        private readonly AssemblyPanelClassifier panelClassifier = new AssemblyPanelClassifier();
        private readonly AssemblyNicheDetector nicheDetector = new AssemblyNicheDetector();

        public AssemblyScene Load(Database db, ObjectId containerId)
        {
            var scene = new AssemblyScene();
            if (db == null || containerId.IsNull)
            {
                scene.Warnings.Add("Контейнер не выбран.");
                return scene;
            }

            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                var containerSolid = tr.GetObject(containerId, OpenMode.ForRead, false) as Solid3d;
                if (containerSolid == null || containerSolid.IsErased)
                {
                    scene.Warnings.Add("Контейнер не найден.");
                    return scene;
                }

                AssemblyContainerData containerData;
                if (!AssemblyContainerStorage.TryRead(containerSolid, tr, out containerData))
                {
                    scene.Warnings.Add("У выбранного тела нет metadata контейнера.");
                    return scene;
                }

                AssemblyContainer container;
                if (!containerAnalyzer.TryAnalyze(containerSolid, containerData, out container))
                {
                    scene.Warnings.Add("Не удалось разобрать габариты контейнера.");
                    return scene;
                }

                scene.Container = container;

                BlockTableRecord modelSpace = (BlockTableRecord)tr.GetObject(SymbolUtilityServices.GetBlockModelSpaceId(db), OpenMode.ForRead);
                foreach (ObjectId id in modelSpace)
                {
                    if (id == containerId)
                    {
                        continue;
                    }

                    var entity = tr.GetObject(id, OpenMode.ForRead, false) as Entity;
                    if (entity == null || entity.IsErased)
                    {
                        continue;
                    }

                    var solid = entity as Solid3d;
                    if (solid != null)
                    {
                        if (!BelongsToAssembly(solid, container, tr))
                        {
                            continue;
                        }

                        SpecificationData specification;
                        SpecificationStorage.TryRead(solid, tr, out specification);
                        AssemblyPartData partData;
                        AssemblyPartStorage.TryRead(solid, tr, out partData);

                        AssemblyPanel panel;
                        if (panelClassifier.TryClassify(solid, container, specification, partData, out panel))
                        {
                            scene.Panels.Add(panel);
                        }
                        else
                        {
                            scene.IgnoredSolids.Add(id);
                        }
                        continue;
                    }

                    var blockRef = entity as BlockReference;
                    if (blockRef != null)
                    {
                        if (!BelongsToAssembly(blockRef, container, tr))
                        {
                            continue;
                        }

                        LoadBlockReferencePanels(blockRef, container, tr, scene);
                    }
                }

                tr.Commit();
            }

            if (scene.Container != null)
            {
                scene.Niches.AddRange(nicheDetector.Detect(scene.Container, scene.Panels));
            }

            if (scene.IgnoredSolids.Count > 0)
            {
                scene.Warnings.Add("Часть деталей не распознана и не участвует в расчете ниш: " + scene.IgnoredSolids.Count.ToString());
            }

            if (scene.Niches.Count == 0 && scene.Container != null)
            {
                scene.Warnings.Add("Свободные ниши не найдены.");
            }

            return scene;
        }

        private void LoadBlockReferencePanels(BlockReference blockRef, AssemblyContainer container, Transaction tr, AssemblyScene scene)
        {
            BlockTableRecord definition = tr.GetObject(blockRef.BlockTableRecord, OpenMode.ForRead, false) as BlockTableRecord;
            if (definition == null)
            {
                return;
            }

            foreach (ObjectId nestedId in definition)
            {
                var sourceSolid = tr.GetObject(nestedId, OpenMode.ForRead, false) as Solid3d;
                if (sourceSolid == null || sourceSolid.IsErased)
                {
                    continue;
                }

                SpecificationData specification;
                SpecificationStorage.TryRead(sourceSolid, tr, out specification);
                AssemblyPartData partData;
                AssemblyPartStorage.TryRead(sourceSolid, tr, out partData);

                using (var transformed = sourceSolid.Clone() as Solid3d)
                {
                    if (transformed == null)
                    {
                        continue;
                    }

                    transformed.TransformBy(blockRef.BlockTransform);
                    AssemblyPanel panel;
                    if (panelClassifier.TryClassify(transformed, container, specification, partData, out panel))
                    {
                        panel.ObjectId = blockRef.ObjectId;
                        panel.Handle = blockRef.Handle.ToString() + ":" + sourceSolid.Handle.ToString();
                        scene.Panels.Add(panel);
                    }
                    else
                    {
                        scene.IgnoredSolids.Add(blockRef.ObjectId);
                    }
                }
            }
        }

        private static bool BelongsToAssembly(Entity entity, AssemblyContainer container, Transaction tr)
        {
            string assemblyNumber = string.Empty;

            SpecificationData specification;
            if (SpecificationStorage.TryRead(entity, tr, out specification))
            {
                assemblyNumber = (specification.AssemblyNumber ?? string.Empty).Trim();
            }

            AssemblyPartData partData;
            if (string.IsNullOrWhiteSpace(assemblyNumber) && AssemblyPartStorage.TryRead(entity, tr, out partData))
            {
                assemblyNumber = (partData.AssemblyNumber ?? string.Empty).Trim();
            }

            if (!string.Equals(assemblyNumber, container.AssemblyNumber, StringComparison.CurrentCultureIgnoreCase))
            {
                return false;
            }

            try
            {
                Point3d center;
                var solid = entity as Solid3d;
                if (solid != null && TryGetBrepCenter(solid, out center))
                {
                }
                else
                {
                    Extents3d extents = entity.GeometricExtents;
                    center = new Point3d(
                        (extents.MinPoint.X + extents.MaxPoint.X) * 0.5,
                        (extents.MinPoint.Y + extents.MaxPoint.Y) * 0.5,
                        (extents.MinPoint.Z + extents.MaxPoint.Z) * 0.5);
                }

                Vector3d delta = center - container.Origin;
                double x = delta.DotProduct(container.FrontAxis);
                double y = delta.DotProduct(container.UpAxis);
                double z = delta.DotProduct(container.DepthAxis);
                return x >= -50.0 && x <= container.Width + 50.0
                    && y >= -50.0 && y <= container.Height + 50.0
                    && z >= -100.0 && z <= container.Depth + 50.0;
            }
            catch
            {
                return false;
            }
        }

        private static bool TryGetBrepCenter(Solid3d solid, out Point3d center)
        {
            center = Point3d.Origin;
            var points = new List<Point3d>();
            try
            {
                using (var brep = new Brep(solid))
                {
                    foreach (BrepVertex vertex in brep.Vertices)
                    {
                        AddUniquePoint(points, vertex.Point);
                    }
                }
            }
            catch
            {
                return false;
            }

            if (points.Count == 0)
            {
                return false;
            }

            double x = 0.0;
            double y = 0.0;
            double z = 0.0;
            foreach (Point3d point in points)
            {
                x += point.X;
                y += point.Y;
                z += point.Z;
            }

            center = new Point3d(x / points.Count, y / points.Count, z / points.Count);
            return true;
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
