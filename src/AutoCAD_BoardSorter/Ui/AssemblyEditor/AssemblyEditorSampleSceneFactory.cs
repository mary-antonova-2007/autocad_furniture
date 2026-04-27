using AutoCAD_BoardSorter.Models;
using Autodesk.AutoCAD.Geometry;

namespace AutoCAD_BoardSorter.Ui
{
    internal static class AssemblyEditorSampleSceneFactory
    {
        public static AssemblyScene Create()
        {
            AssemblyScene scene = new AssemblyScene
            {
                Container = new AssemblyContainer
                {
                    Handle = "SAMPLE",
                    AssemblyNumber = "DEMO-1",
                    FrontFaceKey = "front",
                    FrontAxis = Vector3d.XAxis,
                    UpAxis = Vector3d.ZAxis,
                    DepthAxis = Vector3d.YAxis,
                    Origin = Point3d.Origin,
                    Width = 1800.0,
                    Height = 2100.0,
                    Depth = 600.0
                }
            };

            scene.Panels.Add(new AssemblyPanel
            {
                Handle = "P-LEFT",
                PartRole = "VerticalPanel",
                Material = "ЛДСП 16 мм",
                Thickness = 16.0,
                MinX = 0.0,
                MaxX = 16.0,
                MinY = 0.0,
                MaxY = 2100.0,
                MinDepth = 0.0,
                MaxDepth = 600.0,
                IsRecognized = true
            });

            scene.Panels.Add(new AssemblyPanel
            {
                Handle = "P-RIGHT",
                PartRole = "VerticalPanel",
                Material = "ЛДСП 16 мм",
                Thickness = 16.0,
                MinX = 1784.0,
                MaxX = 1800.0,
                MinY = 0.0,
                MaxY = 2100.0,
                MinDepth = 0.0,
                MaxDepth = 600.0,
                IsRecognized = true
            });

            scene.Panels.Add(new AssemblyPanel
            {
                Handle = "P-BOTTOM",
                PartRole = "HorizontalPanel",
                Material = "ЛДСП 16 мм",
                Thickness = 16.0,
                MinX = 16.0,
                MaxX = 1784.0,
                MinY = 0.0,
                MaxY = 16.0,
                MinDepth = 0.0,
                MaxDepth = 600.0,
                IsRecognized = true
            });

            scene.Panels.Add(new AssemblyPanel
            {
                Handle = "P-TOP",
                PartRole = "HorizontalPanel",
                Material = "ЛДСП 16 мм",
                Thickness = 16.0,
                MinX = 16.0,
                MaxX = 1784.0,
                MinY = 2084.0,
                MaxY = 2100.0,
                MinDepth = 0.0,
                MaxDepth = 600.0,
                IsRecognized = true
            });

            scene.Panels.Add(new AssemblyPanel
            {
                Handle = "P-SHELF",
                PartRole = "HorizontalPanel",
                Material = "ЛДСП 16 мм",
                Thickness = 16.0,
                MinX = 16.0,
                MaxX = 1784.0,
                MinY = 980.0,
                MaxY = 996.0,
                MinDepth = 0.0,
                MaxDepth = 600.0,
                IsRecognized = true
            });

            scene.Panels.Add(new AssemblyPanel
            {
                Handle = "P-BACK-TOP",
                PartRole = "BackPanel",
                Material = "ХДФ 4 мм",
                Thickness = 4.0,
                MinX = 16.0,
                MaxX = 1784.0,
                MinY = 996.0,
                MaxY = 2084.0,
                MinDepth = 596.0,
                MaxDepth = 600.0,
                IsRecognized = true
            });

            scene.Niches.Add(new AssemblyNiche
            {
                Id = "N1",
                MinX = 16.0,
                MaxX = 1784.0,
                MinY = 16.0,
                MaxY = 980.0,
                Depth = 600.0
            });

            scene.Niches.Add(new AssemblyNiche
            {
                Id = "N2",
                MinX = 16.0,
                MaxX = 1784.0,
                MinY = 996.0,
                MaxY = 2084.0,
                Depth = 596.0
            });

            scene.Warnings.Add("Демо-сцена: подключи backend для загрузки из AutoCAD.");
            return scene;
        }
    }
}
