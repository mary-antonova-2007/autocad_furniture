using System;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;
using AutoCAD_BoardSorter.Models;
using AutoCAD_BoardSorter.Ui;

namespace AutoCAD_BoardSorter
{
    public sealed partial class BoardSorterCommands
    {
        [CommandMethod("BDASSEMBLYPALETTE")]
        public void ShowAssemblyPalette()
        {
            AssemblyPalette.Show();
        }

        [CommandMethod("BDMATERIALS")]
        public void ShowMaterialDatabaseEditor()
        {
            var window = new MaterialDatabaseEditorWindow();
            Autodesk.AutoCAD.ApplicationServices.Application.ShowModalWindow(window);
        }

        [CommandMethod("BDMAKECONTAINER")]
        public void MakeAssemblyContainer()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            if (doc == null)
            {
                return;
            }

            Editor ed = doc.Editor;
            Database db = doc.Database;

            string assemblyNumber = PromptString(ed, "\nНомер сборки контейнера", true);
            if (assemblyNumber == null)
            {
                return;
            }

            PromptPointResult basePointResult = ed.GetPoint("\nБазовая точка контейнера: ");
            if (basePointResult.Status != PromptStatus.OK)
            {
                return;
            }

            double width = PromptDouble(ed, "\nШирина контейнера", 1000.0);
            if (double.IsNaN(width) || width <= 0.0)
            {
                return;
            }

            double height = PromptDouble(ed, "\nВысота контейнера", 2000.0);
            if (double.IsNaN(height) || height <= 0.0)
            {
                return;
            }

            double depth = PromptDouble(ed, "\nГлубина контейнера", 600.0);
            if (double.IsNaN(depth) || depth <= 0.0)
            {
                return;
            }

            int axisIndex;
            bool useMax;
            if (!PromptFrontSide(ed, out axisIndex, out useMax))
            {
                return;
            }

            Point3d min = basePointResult.Value;
            double dx = width;
            double dy = depth;
            double dz = height;

            ObjectId containerId;
            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                EnsureLayer(db, tr, AssemblyConstants.ContainerLayer);
                BlockTableRecord modelSpace = (BlockTableRecord)tr.GetObject(SymbolUtilityServices.GetBlockModelSpaceId(db), OpenMode.ForWrite);

                var solid = new Solid3d();
                solid.SetDatabaseDefaults();
                solid.Layer = AssemblyConstants.ContainerLayer;
                solid.CreateBox(dx, dy, dz);
                Point3d boxCenter = min + new Vector3d(dx * 0.5, dy * 0.5, dz * 0.5);
                solid.TransformBy(Matrix3d.Displacement(boxCenter - Point3d.Origin));

                modelSpace.AppendEntity(solid);
                tr.AddNewlyCreatedDBObject(solid, true);
                containerId = solid.ObjectId;
                tr.Commit();
            }

            Vector3d frontAxis;
            Vector3d upAxis;
            Vector3d depthAxis;
            AssemblyContainerAnalyzer.BuildAxesFromFace(axisIndex, useMax, out frontAxis, out upAxis, out depthAxis);

            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                var solid = tr.GetObject(containerId, OpenMode.ForWrite, false) as Solid3d;
                if (solid != null && !solid.IsErased)
                {
                    AssemblyContainerStorage.Write(solid, new AssemblyContainerData
                    {
                        AssemblyNumber = (assemblyNumber ?? string.Empty).Trim(),
                        EntityRole = AssemblyConstants.ContainerRole,
                        Version = "1",
                        FrontFaceKey = AssemblyContainerAnalyzer.CreateFrontFaceKey(axisIndex, useMax),
                        FrontAxisX = AssemblyContainerAnalyzer.FormatDouble(frontAxis.X),
                        FrontAxisY = AssemblyContainerAnalyzer.FormatDouble(frontAxis.Y),
                        FrontAxisZ = AssemblyContainerAnalyzer.FormatDouble(frontAxis.Z),
                        UpAxisX = AssemblyContainerAnalyzer.FormatDouble(upAxis.X),
                        UpAxisY = AssemblyContainerAnalyzer.FormatDouble(upAxis.Y),
                        UpAxisZ = AssemblyContainerAnalyzer.FormatDouble(upAxis.Z),
                        DepthAxisX = AssemblyContainerAnalyzer.FormatDouble(depthAxis.X),
                        DepthAxisY = AssemblyContainerAnalyzer.FormatDouble(depthAxis.Y),
                        DepthAxisZ = AssemblyContainerAnalyzer.FormatDouble(depthAxis.Z),
                        Width = AssemblyContainerAnalyzer.FormatDouble(width),
                        Height = AssemblyContainerAnalyzer.FormatDouble(height),
                        Depth = AssemblyContainerAnalyzer.FormatDouble(depth)
                    }, tr);
                }

                tr.Commit();
            }

            ed.WriteMessage("\nКонтейнер сборки создан.");
        }

        [CommandMethod("BDHIDECONTAINERS")]
        public void HideAssemblyContainers()
        {
            SetContainerLayerVisibility(false);
        }

        [CommandMethod("BDSHOWCONTAINERS")]
        public void ShowAssemblyContainers()
        {
            SetContainerLayerVisibility(true);
        }

        [CommandMethod("BDISOLATEASSEMBLY")]
        public void IsolateAssembly()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            if (doc == null)
            {
                return;
            }

            Database db = doc.Database;
            Editor ed = doc.Editor;
            SetContainerLayerVisibility(true);
            ObjectId containerId = PromptAssemblyContainer(ed, db);
            if (containerId.IsNull)
            {
                return;
            }

            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                var containerSolid = tr.GetObject(containerId, OpenMode.ForRead, false) as Solid3d;
                if (containerSolid == null || containerSolid.IsErased)
                {
                    return;
                }

                AssemblyContainerData containerData;
                if (!AssemblyContainerStorage.TryRead(containerSolid, tr, out containerData))
                {
                    ed.WriteMessage("\nУ выбранного тела нет metadata контейнера.");
                    return;
                }

                string assemblyNumber = (containerData.AssemblyNumber ?? string.Empty).Trim();
                BlockTableRecord modelSpace = (BlockTableRecord)tr.GetObject(SymbolUtilityServices.GetBlockModelSpaceId(db), OpenMode.ForRead);
                foreach (ObjectId id in modelSpace)
                {
                    var entity = tr.GetObject(id, OpenMode.ForWrite, false) as Entity;
                    if (entity == null || entity.IsErased)
                    {
                        continue;
                    }

                    bool isVisible = id == containerId || BelongsToAssembly(entity, assemblyNumber, tr);
                    entity.Visible = isVisible;
                }

                tr.Commit();
            }

            ed.WriteMessage("\nСборка изолирована.");
        }

        [CommandMethod("BDSHOWALL")]
        public void ShowAllEntities()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            if (doc == null)
            {
                return;
            }

            SetContainerLayerVisibility(true);
            Database db = doc.Database;
            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                BlockTableRecord modelSpace = (BlockTableRecord)tr.GetObject(SymbolUtilityServices.GetBlockModelSpaceId(db), OpenMode.ForRead);
                foreach (ObjectId id in modelSpace)
                {
                    var entity = tr.GetObject(id, OpenMode.ForWrite, false) as Entity;
                    if (entity == null || entity.IsErased)
                    {
                        continue;
                    }

                    entity.Visible = true;
                }

                tr.Commit();
            }
        }

        [CommandMethod("BDASSEMBLYEDITOR")]
        public void ShowAssemblyEditor()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            if (doc == null)
            {
                return;
            }

            ObjectId containerId = PromptAssemblyContainer(doc.Editor, doc.Database);
            if (containerId.IsNull)
            {
                return;
            }

            AssemblyEditorWindowHost.Show(containerId);
        }

        private static ObjectId PromptAssemblyContainer(Editor ed, Database db)
        {
            PromptEntityResult result = ed.GetEntity("\nВыбери контейнер сборки или деталь сборки: ");
            if (result.Status != PromptStatus.OK)
            {
                return ObjectId.Null;
            }

            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                var solid = tr.GetObject(result.ObjectId, OpenMode.ForRead, false) as Solid3d;
                if (solid == null || solid.IsErased)
                {
                    return ObjectId.Null;
                }

                AssemblyContainerData containerData;
                if (AssemblyContainerStorage.TryRead(solid, tr, out containerData))
                {
                    tr.Commit();
                    return result.ObjectId;
                }

                string assemblyNumber = string.Empty;
                SpecificationData specification;
                if (SpecificationStorage.TryRead(solid, tr, out specification))
                {
                    assemblyNumber = (specification.AssemblyNumber ?? string.Empty).Trim();
                }

                if (string.IsNullOrWhiteSpace(assemblyNumber))
                {
                    AssemblyPartData partData;
                    if (AssemblyPartStorage.TryRead(solid, tr, out partData))
                    {
                        assemblyNumber = (partData.AssemblyNumber ?? string.Empty).Trim();
                    }
                }

                if (string.IsNullOrWhiteSpace(assemblyNumber))
                {
                    return ObjectId.Null;
                }

                BlockTableRecord modelSpace = (BlockTableRecord)tr.GetObject(SymbolUtilityServices.GetBlockModelSpaceId(db), OpenMode.ForRead);
                foreach (ObjectId id in modelSpace)
                {
                    var candidate = tr.GetObject(id, OpenMode.ForRead, false) as Solid3d;
                    if (candidate == null || candidate.IsErased)
                    {
                        continue;
                    }

                    if (!string.Equals(candidate.Layer, AssemblyConstants.ContainerLayer, StringComparison.CurrentCultureIgnoreCase))
                    {
                        continue;
                    }

                    AssemblyContainerData candidateData;
                    if (AssemblyContainerStorage.TryRead(candidate, tr, out candidateData)
                        && string.Equals((candidateData.AssemblyNumber ?? string.Empty).Trim(), assemblyNumber, StringComparison.CurrentCultureIgnoreCase))
                    {
                        tr.Commit();
                        return id;
                    }
                }

                tr.Commit();
            }

            return ObjectId.Null;
        }

        private static bool BelongsToAssembly(Entity entity, string assemblyNumber, Transaction tr)
        {
            var solid = entity as Solid3d;
            if (solid == null)
            {
                return false;
            }

            SpecificationData specification;
            if (SpecificationStorage.TryRead(solid, tr, out specification)
                && string.Equals((specification.AssemblyNumber ?? string.Empty).Trim(), assemblyNumber, StringComparison.CurrentCultureIgnoreCase))
            {
                return true;
            }

            AssemblyPartData partData;
            return AssemblyPartStorage.TryRead(solid, tr, out partData)
                && string.Equals((partData.AssemblyNumber ?? string.Empty).Trim(), assemblyNumber, StringComparison.CurrentCultureIgnoreCase);
        }

        private static void EnsureLayer(Database db, Transaction tr, string layerName)
        {
            LayerTable layerTable = (LayerTable)tr.GetObject(db.LayerTableId, OpenMode.ForRead);
            if (layerTable.Has(layerName))
            {
                return;
            }

            layerTable.UpgradeOpen();
            var layer = new LayerTableRecord
            {
                Name = layerName
            };
            layerTable.Add(layer);
            tr.AddNewlyCreatedDBObject(layer, true);
        }

        private static void SetContainerLayerVisibility(bool visible)
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            if (doc == null)
            {
                return;
            }

            Database db = doc.Database;
            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                LayerTable layerTable = (LayerTable)tr.GetObject(db.LayerTableId, OpenMode.ForRead);
                if (layerTable.Has(AssemblyConstants.ContainerLayer))
                {
                    LayerTableRecord layer = (LayerTableRecord)tr.GetObject(layerTable[AssemblyConstants.ContainerLayer], OpenMode.ForWrite);
                    layer.IsOff = !visible;
                    layer.IsFrozen = false;
                }

                tr.Commit();
            }
        }

        private static void ResolveNearestFace(Point3d min, Point3d max, Point3d picked, out int axisIndex, out bool useMax)
        {
            double bestDistance = double.MaxValue;
            axisIndex = 1;
            useMax = true;

            TryFace(0, false, Math.Abs(picked.X - min.X), ref bestDistance, ref axisIndex, ref useMax);
            TryFace(0, true, Math.Abs(picked.X - max.X), ref bestDistance, ref axisIndex, ref useMax);
            TryFace(1, false, Math.Abs(picked.Y - min.Y), ref bestDistance, ref axisIndex, ref useMax);
            TryFace(1, true, Math.Abs(picked.Y - max.Y), ref bestDistance, ref axisIndex, ref useMax);
        }

        private static void TryFace(int candidateAxis, bool candidateUseMax, double distance, ref double bestDistance, ref int axisIndex, ref bool useMax)
        {
            if (distance < bestDistance)
            {
                bestDistance = distance;
                axisIndex = candidateAxis;
                useMax = candidateUseMax;
            }
        }

        private static bool PromptFrontSide(Editor ed, out int axisIndex, out bool useMax)
        {
            axisIndex = 1;
            useMax = false;

            var options = new PromptKeywordOptions("\nФронтальная сторона [Спереди/Сзади/Слева/Справа/Снизу/Сверху] <Спереди>: ")
            {
                AllowNone = true
            };
            options.Keywords.Add("Спереди");
            options.Keywords.Add("Сзади");
            options.Keywords.Add("Слева");
            options.Keywords.Add("Справа");
            options.Keywords.Add("Снизу");
            options.Keywords.Add("Сверху");
            options.Keywords.Default = "Спереди";

            PromptResult result = ed.GetKeywords(options);
            if (result.Status == PromptStatus.Cancel)
            {
                return false;
            }

            string side = result.Status == PromptStatus.OK ? result.StringResult : "Спереди";
            switch (side)
            {
                case "Слева":
                    axisIndex = 0;
                    useMax = false;
                    return true;
                case "Справа":
                    axisIndex = 0;
                    useMax = true;
                    return true;
                case "Сзади":
                    axisIndex = 1;
                    useMax = false;
                    return true;
                case "Снизу":
                    axisIndex = 2;
                    useMax = false;
                    return true;
                case "Сверху":
                    axisIndex = 2;
                    useMax = true;
                    return true;
                default:
                    axisIndex = 1;
                    useMax = true;
                    return true;
            }
        }
    }
}
