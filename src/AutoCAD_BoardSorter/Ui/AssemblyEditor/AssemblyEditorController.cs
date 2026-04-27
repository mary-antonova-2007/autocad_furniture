using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Windows;
using AutoCAD_BoardSorter.Models;

namespace AutoCAD_BoardSorter.Ui
{
    internal sealed class AssemblyEditorController
    {
        private readonly IAssemblyEditorBackend backend;
        private readonly AssemblyEditorProjectionBuilder projectionBuilder = new AssemblyEditorProjectionBuilder();
        private readonly AssemblyEditorHitTestService hitTestService = new AssemblyEditorHitTestService();
        private readonly AssemblyEditorPreviewBuilder previewBuilder = new AssemblyEditorPreviewBuilder();
        private Size viewportSize;
        private AssemblyEditorRenderModel renderModel;
        private AssemblyEditorPreviewVisual preview;

        public AssemblyEditorController(IAssemblyEditorBackend backend)
        {
            this.backend = backend ?? new AssemblyEditorLocalBackend();
            State = new AssemblyEditorState
            {
                Tool = AssemblyEditorTool.Select,
                ActiveMaterial = "ЛДСП 16 мм",
                ActiveThickness = 16.0,
                ActiveOffset = 0.0,
                ViewZoom = 1.0,
                StatusText = "Выбери контейнер или загрузи сцену.",
                DrawerSettings = CreateDefaultDrawerSettings()
            };
        }

        public AssemblyEditorState State { get; }

        public AssemblyEditorRenderModel RenderModel
        {
            get { return renderModel; }
        }

        public void LoadScene(AssemblyScene scene)
        {
            State.Scene = scene ?? AssemblyEditorSampleSceneFactory.Create();
            if (State.Scene.Niches.Count == 0 && State.Scene.Container != null)
            {
                State.Scene = backend.Reload(State.Scene);
            }

            State.HoveredNiche = null;
            State.SelectedNiche = State.Scene.Niches.FirstOrDefault();
            UpdateStatusText();
            Rebuild();
        }

        public void SetTool(AssemblyEditorTool tool)
        {
            State.Tool = tool;
            RebuildPreviewFromSelection();
            UpdateStatusText();
            Rebuild();
        }

        public void SetMaterial(string material)
        {
            State.ActiveMaterial = string.IsNullOrWhiteSpace(material) ? string.Empty : material.Trim();
            UpdateStatusText();
        }

        public void SetThickness(double thickness)
        {
            State.ActiveThickness = Math.Max(1.0, thickness);
            RebuildPreviewFromSelection();
            UpdateStatusText();
            Rebuild();
        }

        public void SetOffset(double offset)
        {
            State.ActiveOffset = Math.Max(0.0, offset);
            RebuildPreviewFromSelection();
            UpdateStatusText();
            Rebuild();
        }

        public void UpdateDrawerSettings(Action<AssemblyDrawerSettings> update)
        {
            if (State.DrawerSettings == null)
            {
                State.DrawerSettings = CreateDefaultDrawerSettings();
            }

            if (update != null)
            {
                update(State.DrawerSettings);
            }

            AssemblyDrawerLayoutCalculator.EnsureDrafts(State.DrawerSettings);
            RebuildPreviewFromSelection();
            UpdateStatusText();
            Rebuild();
        }

        public void SetViewportSize(Size size)
        {
            viewportSize = size;
            Rebuild();
        }

        public void SetShiftPressed(bool value)
        {
            State.IsShiftPressed = value;
            RebuildPreviewFromSelection();
            UpdateStatusText();
            Rebuild();
        }

        public void ZoomAt(Point screenPoint, double deltaScale)
        {
            if (renderModel == null || renderModel.Projection == null)
            {
                return;
            }

            double currentZoom = State.ViewZoom <= 0.0 ? 1.0 : State.ViewZoom;
            double nextZoom = Math.Max(0.2, Math.Min(8.0, currentZoom * deltaScale));
            if (Math.Abs(nextZoom - currentZoom) <= 1e-6)
            {
                return;
            }

            Point modelBefore = renderModel.Projection.ScreenToModel(screenPoint);
            State.ViewZoom = nextZoom;
            Rebuild();
            if (renderModel == null || renderModel.Projection == null)
            {
                return;
            }

            Point screenAfter = renderModel.Projection.ModelToScreen(modelBefore.X, modelBefore.Y);
            State.ViewPanX += screenPoint.X - screenAfter.X;
            State.ViewPanY += screenPoint.Y - screenAfter.Y;
            Rebuild();
        }

        public void Pan(double deltaX, double deltaY)
        {
            State.ViewPanX += deltaX;
            State.ViewPanY += deltaY;
            Rebuild();
        }

        public void ResetView()
        {
            State.ViewZoom = 1.0;
            State.ViewPanX = 0.0;
            State.ViewPanY = 0.0;
            Rebuild();
        }

        public void HandlePointerMove(Point screenPoint)
        {
            if (renderModel == null || renderModel.Projection == null)
            {
                return;
            }

            AssemblyEditorNicheVisual hit = hitTestService.HitNiche(renderModel, screenPoint);
            State.HoveredNiche = hit != null ? hit.Source : null;

            Point modelPoint = renderModel.Projection.ScreenToModel(screenPoint);
            AssemblyNiche niche = State.SelectedNiche ?? State.HoveredNiche;
            if (State.HoveredNiche != null)
            {
                niche = State.HoveredNiche;
            }

            preview = previewBuilder.Build(State, renderModel.Projection, niche, modelPoint);
            UpdateStatusText();
            Rebuild();
        }

        public void HandlePointerLeave()
        {
            State.HoveredNiche = null;
            preview = previewBuilder.Build(State, renderModel != null ? renderModel.Projection : null, State.SelectedNiche, GetSelectedNicheCenter());
            UpdateStatusText();
            Rebuild();
        }

        public bool HandleClick(Point screenPoint)
        {
            if (renderModel == null)
            {
                PaletteDebugLogger.Info(null, "AssemblyEditorController HandleClick skipped: renderModel is null");
                return false;
            }

            if (State.Tool != AssemblyEditorTool.Select && preview != null)
            {
                if (State.HoveredNiche != null)
                {
                    State.SelectedNiche = State.HoveredNiche;
                }

                Point modelPoint = renderModel.Projection != null
                    ? renderModel.Projection.ScreenToModel(screenPoint)
                    : new Point();
                preview = previewBuilder.Build(State, renderModel.Projection, State.SelectedNiche, modelPoint) ?? preview;
                PaletteDebugLogger.Info(
                    null,
                    "AssemblyEditorController HandleClick using active preview niche="
                    + (State.SelectedNiche != null ? State.SelectedNiche.Id : "<null>")
                    + " hovered="
                    + (State.HoveredNiche != null ? State.HoveredNiche.Id : "<null>"));
                UpdateStatusText();
                Rebuild();
                return State.SelectedNiche != null;
            }

            AssemblyEditorNicheVisual hit = hitTestService.HitNiche(renderModel, screenPoint);
            if (hit != null)
            {
                PaletteDebugLogger.Info(null, "AssemblyEditorController HandleClick niche=" + hit.Source.Id);
                State.SelectedNiche = hit.Source;
                Point modelPoint = renderModel.Projection.ScreenToModel(screenPoint);
                preview = previewBuilder.Build(State, renderModel.Projection, State.SelectedNiche, modelPoint);
                UpdateStatusText();
                Rebuild();
                return true;
            }

            PaletteDebugLogger.Info(null, "AssemblyEditorController HandleClick: niche not hit");
            return false;
        }

        public bool CommitPreview()
        {
            if (preview == null || State.Scene == null || State.SelectedNiche == null)
            {
                PaletteDebugLogger.Info(
                    null,
                    "AssemblyEditorController CommitPreview skipped: preview="
                    + (preview != null) + " scene=" + (State.Scene != null) + " selectedNiche=" + (State.SelectedNiche != null));
                return false;
            }

            AssemblyEditorInsertRequest request = new AssemblyEditorInsertRequest
            {
                Scene = State.Scene,
                Niche = State.SelectedNiche,
                Tool = State.Tool,
                Material = State.ActiveMaterial,
                Thickness = State.ActiveThickness,
                ModelRect = preview.ModelBounds,
                ModelRects = preview.ModelRects,
                IsFlush = preview.IsFlush,
                FlushSide = preview.FlushSide,
                AvailableDepth = preview.AvailableDepth,
                DepthOffset = preview.DepthOffset,
                DrawerSegments = preview.DrawerSegments
            };

            PaletteDebugLogger.Info(
                null,
                "AssemblyEditorController CommitPreview niche=" + request.Niche.Id
                + " tool=" + request.Tool
                + " rect=" + request.ModelRect.ToString());
            State.Scene = backend.ApplyInsert(request) ?? State.Scene;
            State.SelectedNiche = PickClosestNiche(State.Scene.Niches, request.Niche);
            State.HoveredNiche = null;
            RebuildPreviewFromSelection();
            UpdateStatusText();
            Rebuild();
            return true;
        }

        public void Reload()
        {
            if (State.Scene == null)
            {
                return;
            }

            State.Scene = backend.Reload(State.Scene) ?? State.Scene;
            State.SelectedNiche = PickClosestNiche(State.Scene.Niches, State.SelectedNiche);
            RebuildPreviewFromSelection();
            UpdateStatusText();
            Rebuild();
        }

        private void Rebuild()
        {
            renderModel = projectionBuilder.Build(State, viewportSize, preview);
        }

        private void RebuildPreviewFromSelection()
        {
            if (renderModel == null || renderModel.Projection == null)
            {
                preview = null;
                return;
            }

            preview = previewBuilder.Build(State, renderModel.Projection, State.SelectedNiche, GetSelectedNicheCenter());
        }

        private Point GetSelectedNicheCenter()
        {
            AssemblyNiche niche = State.SelectedNiche;
            if (niche == null)
            {
                return new Point();
            }

            return new Point((niche.MinX + niche.MaxX) * 0.5, (niche.MinY + niche.MaxY) * 0.5);
        }

        private void UpdateStatusText()
        {
            if (State.Scene == null || State.Scene.Container == null)
            {
                State.StatusText = "Сцена не загружена.";
                return;
            }

            List<string> chunks = new List<string>();
            chunks.Add("Сборка: " + Safe(State.Scene.Container.AssemblyNumber));

            if (State.SelectedNiche != null)
            {
                chunks.Add(string.Format(
                    CultureInfo.InvariantCulture,
                    "Ниша {0:0.#} x {1:0.#} x {2:0.#}",
                    State.SelectedNiche.Width,
                    State.SelectedNiche.Height,
                    State.SelectedNiche.Depth));
            }

            if (preview != null && State.Tool != AssemblyEditorTool.Select)
            {
                chunks.Add(preview.Label.Replace("\n", " | "));
            }
            else
            {
                chunks.Add("Инструмент: " + ResolveToolTitle(State.Tool));
            }

            if (State.Scene.Warnings.Count > 0)
            {
                chunks.Add("Предупреждений: " + State.Scene.Warnings.Count.ToString(CultureInfo.InvariantCulture));
            }

            State.StatusText = string.Join("  •  ", chunks.Where(chunk => !string.IsNullOrWhiteSpace(chunk)));
        }

        private static string ResolveToolTitle(AssemblyEditorTool tool)
        {
            switch (tool)
            {
                case AssemblyEditorTool.VerticalPanel:
                    return "Стойка";
                case AssemblyEditorTool.HorizontalPanel:
                    return "Полка";
                case AssemblyEditorTool.BackPanel:
                    return "Задняя стенка";
                case AssemblyEditorTool.FrontPanel:
                    return "Передняя стенка";
                case AssemblyEditorTool.Drawers:
                    return "Ящики";
                default:
                    return "Выбор";
            }
        }

        private static AssemblyDrawerSettings CreateDefaultDrawerSettings()
        {
            var settings = new AssemblyDrawerSettings
            {
                Count = 3,
                Mode = AssemblyDrawerMode.Overlay,
                GapLeft = 14.5,
                GapRight = 14.5,
                GapTop = 14.5,
                GapBottom = 14.5,
                GapBetween = 3.0,
                FrontGap = 3.0,
                AutoDepth = true,
                Depth = 450.0,
                FrontMaterial = "Фасад",
                BodyMaterial = "ЛДСП 16 мм",
                BottomMaterial = "Дно 16 мм",
                FrontThickness = 18.0,
                BottomThickness = 16.0
            };
            AssemblyDrawerLayoutCalculator.EnsureDrafts(settings);
            return settings;
        }

        private static string Safe(string value)
        {
            return string.IsNullOrWhiteSpace(value) ? "-" : value.Trim();
        }

        private static AssemblyNiche PickClosestNiche(IEnumerable<AssemblyNiche> niches, AssemblyNiche previous)
        {
            if (niches == null)
            {
                return null;
            }

            List<AssemblyNiche> nicheList = niches.ToList();
            if (nicheList.Count == 0)
            {
                return null;
            }

            if (previous == null)
            {
                return nicheList[0];
            }

            double targetX = (previous.MinX + previous.MaxX) * 0.5;
            double targetY = (previous.MinY + previous.MaxY) * 0.5;

            return nicheList
                .OrderBy(niche =>
                {
                    double centerX = (niche.MinX + niche.MaxX) * 0.5;
                    double centerY = (niche.MinY + niche.MaxY) * 0.5;
                    double dx = centerX - targetX;
                    double dy = centerY - targetY;
                    return (dx * dx) + (dy * dy);
                })
                .First();
        }

        private sealed class AssemblyEditorLocalBackend : IAssemblyEditorBackend
        {
            private const double MinimumNicheSize = 30.0;

            public AssemblyScene Reload(AssemblyScene scene)
            {
                AssemblyScene copy = CloneScene(scene);
                RebuildNiches(copy);
                return copy;
            }

            public AssemblyScene ApplyInsert(AssemblyEditorInsertRequest request)
            {
                AssemblyScene scene = CloneScene(request != null ? request.Scene : null);
                if (scene == null || request == null)
                {
                    return scene;
                }

                Rect modelRect = request.ModelRect;
                if (request.Tool == AssemblyEditorTool.Drawers && request.DrawerSegments != null)
                {
                    foreach (AssemblyDrawerSegment segment in request.DrawerSegments)
                    {
                        scene.Panels.Add(new AssemblyPanel
                        {
                            Handle = "UI-" + Guid.NewGuid().ToString("N"),
                            PartRole = segment.Role,
                            Material = segment.Material,
                            Thickness = Math.Min(segment.MaxX - segment.MinX, Math.Min(segment.MaxY - segment.MinY, segment.MaxZ - segment.MinZ)),
                            Length = Math.Max(segment.MaxX - segment.MinX, Math.Max(segment.MaxY - segment.MinY, segment.MaxZ - segment.MinZ)),
                            Width = (segment.MaxX - segment.MinX) + (segment.MaxY - segment.MinY) + (segment.MaxZ - segment.MinZ)
                                - Math.Min(segment.MaxX - segment.MinX, Math.Min(segment.MaxY - segment.MinY, segment.MaxZ - segment.MinZ))
                                - Math.Max(segment.MaxX - segment.MinX, Math.Max(segment.MaxY - segment.MinY, segment.MaxZ - segment.MinZ)),
                            MinX = segment.MinX,
                            MaxX = segment.MaxX,
                            MinY = segment.MinY,
                            MaxY = segment.MaxY,
                            MinDepth = segment.MinZ,
                            MaxDepth = segment.MaxZ,
                            IsRecognized = true
                        });
                    }
                }
                else
                {
                    scene.Panels.Add(new AssemblyPanel
                {
                    Handle = "UI-" + Guid.NewGuid().ToString("N"),
                    PartRole = ResolveRole(request.Tool),
                    Material = request.Material,
                    Thickness = request.Thickness,
                    Length = string.Equals(ResolveRole(request.Tool), "VerticalPanel", StringComparison.OrdinalIgnoreCase)
                        ? modelRect.Height
                        : modelRect.Width,
                    Width = request.AvailableDepth,
                    MinX = modelRect.Left,
                    MaxX = modelRect.Right,
                    MinY = modelRect.Top,
                    MaxY = modelRect.Bottom,
                    MinDepth = string.Equals(ResolveRole(request.Tool), "BackPanel", StringComparison.OrdinalIgnoreCase)
                        ? Math.Max(0.0, request.AvailableDepth - request.Thickness)
                        : 0.0,
                    MaxDepth = request.AvailableDepth,
                    IsRecognized = true
                });
                }

                RebuildNiches(scene);
                return scene;
            }

            private static AssemblyScene CloneScene(AssemblyScene source)
            {
                if (source == null)
                {
                    return null;
                }

                AssemblyScene copy = new AssemblyScene
                {
                    Container = source.Container == null
                        ? null
                        : new AssemblyContainer
                        {
                            ObjectId = source.Container.ObjectId,
                            Handle = source.Container.Handle,
                            AssemblyNumber = source.Container.AssemblyNumber,
                            FrontFaceKey = source.Container.FrontFaceKey,
                            FrontAxis = source.Container.FrontAxis,
                            UpAxis = source.Container.UpAxis,
                            DepthAxis = source.Container.DepthAxis,
                            Bounds = source.Container.Bounds,
                            Origin = source.Container.Origin,
                            Width = source.Container.Width,
                            Height = source.Container.Height,
                            Depth = source.Container.Depth
                        }
                };

                foreach (AssemblyPanel panel in source.Panels)
                {
                    copy.Panels.Add(new AssemblyPanel
                    {
                        ObjectId = panel.ObjectId,
                        Handle = panel.Handle,
                        PartRole = panel.PartRole,
                        Material = panel.Material,
                        Thickness = panel.Thickness,
                        Length = panel.Length,
                        Width = panel.Width,
                        MinX = panel.MinX,
                        MaxX = panel.MaxX,
                        MinY = panel.MinY,
                        MaxY = panel.MaxY,
                        MinDepth = panel.MinDepth,
                        MaxDepth = panel.MaxDepth,
                        IsRecognized = panel.IsRecognized
                    });
                }

                foreach (var ignored in source.IgnoredSolids)
                {
                    copy.IgnoredSolids.Add(ignored);
                }

                foreach (string warning in source.Warnings)
                {
                    copy.Warnings.Add(warning);
                }

                foreach (AssemblyNiche niche in source.Niches)
                {
                    copy.Niches.Add(new AssemblyNiche
                    {
                        Id = niche.Id,
                        MinX = niche.MinX,
                        MaxX = niche.MaxX,
                        MinY = niche.MinY,
                        MaxY = niche.MaxY,
                        Depth = niche.Depth
                    });
                }

                return copy;
            }

            private static void RebuildNiches(AssemblyScene scene)
            {
                scene.Niches.Clear();
                if (scene.Container == null)
                {
                    return;
                }

                List<double> xs = new List<double> { 0.0, scene.Container.Width };
                List<double> ys = new List<double> { 0.0, scene.Container.Height };

                foreach (AssemblyPanel panel in scene.Panels.Where(panel => panel.IsRecognized))
                {
                    xs.Add(panel.MinX);
                    xs.Add(panel.MaxX);
                    ys.Add(panel.MinY);
                    ys.Add(panel.MaxY);
                }

                xs = xs.Distinct().OrderBy(value => value).ToList();
                ys = ys.Distinct().OrderBy(value => value).ToList();

                int columns = Math.Max(0, xs.Count - 1);
                int rows = Math.Max(0, ys.Count - 1);
                bool[,] free = new bool[columns, rows];
                double[,] depths = new double[columns, rows];

                for (int ix = 0; ix < columns; ix++)
                {
                    for (int iy = 0; iy < rows; iy++)
                    {
                        double minX = xs[ix];
                        double maxX = xs[ix + 1];
                        double minY = ys[iy];
                        double maxY = ys[iy + 1];
                        double centerX = (minX + maxX) * 0.5;
                        double centerY = (minY + maxY) * 0.5;

                        bool occupied = scene.Panels.Any(panel =>
                            panel.IsRecognized
                            && centerX > panel.MinX
                            && centerX < panel.MaxX
                            && centerY > panel.MinY
                            && centerY < panel.MaxY);

                        free[ix, iy] = !occupied;
                        depths[ix, iy] = ResolveDepth(scene, centerX, centerY);
                    }
                }

                bool[,] used = new bool[columns, rows];
                int nicheIndex = 1;

                for (int ix = 0; ix < columns; ix++)
                {
                    for (int iy = 0; iy < rows; iy++)
                    {
                        if (!free[ix, iy] || used[ix, iy])
                        {
                            continue;
                        }

                        int maxIx = ix;
                        while (maxIx + 1 < columns && free[maxIx + 1, iy] && !used[maxIx + 1, iy] && NearlyEqual(depths[maxIx + 1, iy], depths[ix, iy]))
                        {
                            maxIx++;
                        }

                        int maxIy = iy;
                        bool canExtend = true;
                        while (canExtend && maxIy + 1 < rows)
                        {
                            for (int scanX = ix; scanX <= maxIx; scanX++)
                            {
                                if (!free[scanX, maxIy + 1] || used[scanX, maxIy + 1] || !NearlyEqual(depths[scanX, maxIy + 1], depths[ix, iy]))
                                {
                                    canExtend = false;
                                    break;
                                }
                            }

                            if (canExtend)
                            {
                                maxIy++;
                            }
                        }

                        for (int fillX = ix; fillX <= maxIx; fillX++)
                        {
                            for (int fillY = iy; fillY <= maxIy; fillY++)
                            {
                                used[fillX, fillY] = true;
                            }
                        }

                        scene.Niches.Add(new AssemblyNiche
                        {
                            Id = "N" + nicheIndex.ToString(CultureInfo.InvariantCulture),
                            MinX = xs[ix],
                            MaxX = xs[maxIx + 1],
                            MinY = ys[iy],
                            MaxY = ys[maxIy + 1],
                            Depth = depths[ix, iy]
                        });
                        if (scene.Niches[scene.Niches.Count - 1].Width <= MinimumNicheSize
                            || scene.Niches[scene.Niches.Count - 1].Height <= MinimumNicheSize)
                        {
                            scene.Niches.RemoveAt(scene.Niches.Count - 1);
                            continue;
                        }

                        nicheIndex++;
                    }
                }

                if (scene.IgnoredSolids.Count > 0 && !scene.Warnings.Any())
                {
                    scene.Warnings.Add("Часть тел не распознана и не участвует в расчете ниш.");
                }
            }

            private static double ResolveDepth(AssemblyScene scene, double centerX, double centerY)
            {
                double depth = scene.Container != null ? scene.Container.Depth : 0.0;
                foreach (AssemblyPanel panel in scene.Panels)
                {
                    if (!panel.IsRecognized || !string.Equals(panel.PartRole, "BackPanel", StringComparison.OrdinalIgnoreCase))
                    {
                        continue;
                    }

                    if (centerX > panel.MinX && centerX < panel.MaxX && centerY > panel.MinY && centerY < panel.MaxY)
                    {
                        depth = Math.Min(depth, Math.Max(0.0, panel.MinDepth));
                    }
                }

                return depth;
            }

            private static bool NearlyEqual(double left, double right)
            {
                return Math.Abs(left - right) <= 0.001;
            }

            private static string ResolveRole(AssemblyEditorTool tool)
            {
                switch (tool)
                {
                    case AssemblyEditorTool.VerticalPanel:
                        return "VerticalPanel";
                    case AssemblyEditorTool.HorizontalPanel:
                        return "HorizontalPanel";
                    case AssemblyEditorTool.BackPanel:
                        return "BackPanel";
                    case AssemblyEditorTool.FrontPanel:
                        return "FrontPanel";
                    default:
                        return "Unknown";
                }
            }
        }
    }
}
