using System;
using System.Windows;
using Autodesk.AutoCAD.DatabaseServices;
using AutoCAD_BoardSorter.Models;
using AcadApp = Autodesk.AutoCAD.ApplicationServices.Application;

namespace AutoCAD_BoardSorter.Ui
{
    internal sealed class AutoCadAssemblyEditorBackend : IAssemblyEditorBackend
    {
        private readonly ObjectId containerId;
        private readonly AssemblySceneLoader loader = new AssemblySceneLoader();
        private readonly AssemblyPanelGenerator generator = new AssemblyPanelGenerator();
        private readonly AssemblyDrawerGenerator drawerGenerator = new AssemblyDrawerGenerator();

        public AutoCadAssemblyEditorBackend(ObjectId containerId)
        {
            this.containerId = containerId;
        }

        public AssemblyScene Reload(AssemblyScene scene)
        {
            var doc = AcadApp.DocumentManager.MdiActiveDocument;
            Database db = doc == null ? null : doc.Database;
            PaletteDebugLogger.Info(db, "AssemblyEditor Reload start containerId=" + containerId.Handle);
            return loader.Load(db, containerId);
        }

        public AssemblyScene ApplyInsert(AssemblyEditorInsertRequest request)
        {
            if (request == null || request.Tool == AssemblyEditorTool.Select)
            {
                return Reload(request == null ? null : request.Scene);
            }

            var doc = AcadApp.DocumentManager.MdiActiveDocument;
            Database db = doc == null ? null : doc.Database;
            AssemblyScene scene = request.Scene ?? Reload(null);
            if (db == null || scene == null || scene.Container == null || request.Niche == null)
            {
                PaletteDebugLogger.Info(db, "AssemblyEditor ApplyInsert skipped: db/scene/container/niche missing");
                return scene;
            }

            bool anchorToMinSide;
            double offset;
            ResolvePlacement(request, out anchorToMinSide, out offset);
            PaletteDebugLogger.Info(
                db,
                "AssemblyEditor ApplyInsert start tool=" + request.Tool
                + " niche=" + request.Niche.Id
                + " anchorToMinSide=" + anchorToMinSide
                + " offset=" + offset.ToString("0.###", System.Globalization.CultureInfo.InvariantCulture)
                + " thickness=" + request.Thickness.ToString("0.###", System.Globalization.CultureInfo.InvariantCulture)
                + " material=\"" + (request.Material ?? string.Empty) + "\"");

            try
            {
                using (doc.LockDocument())
                {
                    if (request.Tool == AssemblyEditorTool.Drawers)
                    {
                        var layout = new AssemblyDrawerLayout();
                        if (request.DrawerSegments != null)
                        {
                            foreach (AssemblyDrawerSegment segment in request.DrawerSegments)
                            {
                                layout.Segments.Add(segment);
                            }
                        }

                        drawerGenerator.CreateDrawers(db, scene.Container, layout);
                    }
                    else
                    {
                        var rects = request.ModelRects != null && request.ModelRects.Count > 0
                            ? request.ModelRects
                            : new[] { request.ModelRect };
                        foreach (Rect rect in rects)
                        {
                            generator.CreatePanel(
                                db,
                                scene.Container,
                                rect,
                                request.AvailableDepth,
                                request.DepthOffset,
                                request.Tool,
                                request.Thickness,
                                request.Material);
                        }
                    }
                }

                PaletteDebugLogger.Info(db, "AssemblyEditor ApplyInsert success");
            }
            catch (Exception ex)
            {
                PaletteDebugLogger.Error(db, "AssemblyEditor ApplyInsert failed", ex);
                if (doc != null)
                {
                    doc.Editor.WriteMessage("\nОшибка вставки детали конструктора: {0}", ex.Message);
                }

                return scene;
            }

            return Reload(scene);
        }

        private static void ResolvePlacement(AssemblyEditorInsertRequest request, out bool anchorToMinSide, out double offset)
        {
            anchorToMinSide = true;
            offset = 0.0;
            if (request == null || request.Niche == null)
            {
                return;
            }

            Rect rect = request.ModelRect;
            AssemblyNiche niche = request.Niche;
            if (request.Tool == AssemblyEditorTool.VerticalPanel)
            {
                double leftOffset = Math.Max(0.0, rect.Left - niche.MinX);
                double rightOffset = Math.Max(0.0, niche.MaxX - rect.Right);
                anchorToMinSide = leftOffset <= rightOffset;
                offset = anchorToMinSide ? leftOffset : rightOffset;
                return;
            }

            if (request.Tool == AssemblyEditorTool.HorizontalPanel)
            {
                double bottomOffset = Math.Max(0.0, rect.Top - niche.MinY);
                double topOffset = Math.Max(0.0, niche.MaxY - rect.Bottom);
                anchorToMinSide = bottomOffset <= topOffset;
                offset = anchorToMinSide ? bottomOffset : topOffset;
            }
        }
    }
}
