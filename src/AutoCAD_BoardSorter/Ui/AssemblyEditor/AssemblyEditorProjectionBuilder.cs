using System;
using System.Globalization;
using System.Windows;
using System.Windows.Media;
using AutoCAD_BoardSorter.Models;

namespace AutoCAD_BoardSorter.Ui
{
    internal sealed class AssemblyEditorProjectionBuilder
    {
        private const double Padding = 24.0;
        private static readonly Brush ContainerStroke = new SolidColorBrush(Color.FromRgb(120, 140, 166));
        private static readonly Brush VerticalFill = new SolidColorBrush(Color.FromArgb(180, 83, 150, 244));
        private static readonly Brush HorizontalFill = new SolidColorBrush(Color.FromArgb(180, 16, 185, 129));
        private static readonly Brush BackFill = new SolidColorBrush(Color.FromArgb(180, 245, 158, 11));
        private static readonly Brush UnknownFill = new SolidColorBrush(Color.FromArgb(160, 107, 114, 128));
        private static readonly Brush PanelStroke = new SolidColorBrush(Color.FromRgb(226, 232, 240));
        private static readonly Brush NicheFill = new SolidColorBrush(Color.FromArgb(22, 255, 255, 255));
        private static readonly Brush NicheStroke = new SolidColorBrush(Color.FromRgb(71, 85, 105));
        private static readonly Brush HoverFill = new SolidColorBrush(Color.FromArgb(65, 59, 130, 246));
        private static readonly Brush HoverStroke = new SolidColorBrush(Color.FromRgb(96, 165, 250));
        private static readonly Brush SelectedFill = new SolidColorBrush(Color.FromArgb(80, 14, 165, 233));
        private static readonly Brush SelectedStroke = new SolidColorBrush(Color.FromRgb(56, 189, 248));

        public AssemblyEditorRenderModel Build(AssemblyEditorState state, Size viewportSize, AssemblyEditorPreviewVisual preview)
        {
            AssemblyEditorRenderModel render = new AssemblyEditorRenderModel();
            if (state == null || state.Scene == null || state.Scene.Container == null || viewportSize.Width < 1 || viewportSize.Height < 1)
            {
                return render;
            }

            AssemblyContainer container = state.Scene.Container;
            double width = Math.Max(container.Width, 1.0);
            double height = Math.Max(container.Height, 1.0);
            double usableWidth = Math.Max(80.0, viewportSize.Width - (Padding * 2.0));
            double usableHeight = Math.Max(80.0, viewportSize.Height - (Padding * 2.0));
            double baseScale = Math.Min(usableWidth / width, usableHeight / height);
            double scale = baseScale * Math.Max(0.1, state.ViewZoom <= 0.0 ? 1.0 : state.ViewZoom);
            double scaledWidth = width * scale;
            double scaledHeight = height * scale;
            double offsetX = ((viewportSize.Width - scaledWidth) * 0.5) + state.ViewPanX;
            double offsetY = ((viewportSize.Height - scaledHeight) * 0.5) + state.ViewPanY;

            AssemblyEditorProjection projection = new AssemblyEditorProjection
            {
                ViewportBounds = new Rect(0, 0, viewportSize.Width, viewportSize.Height),
                Scale = scale,
                OffsetX = offsetX,
                OffsetY = offsetY,
                MinX = 0.0,
                MaxY = container.Height
            };

            render.Projection = projection;
            render.ContainerRect = projection.ModelRectToScreen(0.0, 0.0, container.Width, container.Height);

            foreach (AssemblyPanel panel in state.Scene.Panels)
            {
                Rect rect = projection.ModelRectToScreen(panel.MinX, panel.MinY, panel.MaxX, panel.MaxY);
                render.Panels.Add(new AssemblyEditorPanelVisual
                {
                    Source = panel,
                    Bounds = rect,
                    Fill = ResolvePanelFill(panel),
                    Stroke = PanelStroke,
                    Label = BuildPanelLabel(panel)
                });
            }

            foreach (AssemblyNiche niche in state.Scene.Niches)
            {
                bool isHovered = ReferenceEquals(niche, state.HoveredNiche);
                bool isSelected = ReferenceEquals(niche, state.SelectedNiche);
                render.Niches.Add(new AssemblyEditorNicheVisual
                {
                    Source = niche,
                    Bounds = projection.ModelRectToScreen(niche.MinX, niche.MinY, niche.MaxX, niche.MaxY),
                    Fill = isSelected ? SelectedFill : (isHovered ? HoverFill : NicheFill),
                    Stroke = isSelected ? SelectedStroke : (isHovered ? HoverStroke : NicheStroke),
                    IsHovered = isHovered,
                    IsSelected = isSelected,
                    Label = string.Format(
                        CultureInfo.InvariantCulture,
                        "{0:0.#} x {1:0.#}\nГл. {2:0.#}",
                        niche.Width,
                        niche.Height,
                        niche.Depth)
                });
            }

            render.Preview = preview;
            return render;
        }

        private static Brush ResolvePanelFill(AssemblyPanel panel)
        {
            string role = panel != null ? panel.PartRole : string.Empty;
            if (string.Equals(role, "VerticalPanel", StringComparison.OrdinalIgnoreCase))
            {
                return VerticalFill;
            }

            if (string.Equals(role, "HorizontalPanel", StringComparison.OrdinalIgnoreCase))
            {
                return HorizontalFill;
            }

            if (string.Equals(role, "BackPanel", StringComparison.OrdinalIgnoreCase))
            {
                return BackFill;
            }

            if (string.Equals(role, "FrontPanel", StringComparison.OrdinalIgnoreCase))
            {
                return new SolidColorBrush(Color.FromArgb(180, 239, 68, 68));
            }

            if (string.Equals(role, AssemblyConstants.DrawerFrontRole, StringComparison.OrdinalIgnoreCase))
            {
                return new SolidColorBrush(Color.FromArgb(190, 168, 85, 247));
            }

            if (string.Equals(role, AssemblyConstants.DrawerBottomRole, StringComparison.OrdinalIgnoreCase)
                || string.Equals(role, AssemblyConstants.DrawerSideRole, StringComparison.OrdinalIgnoreCase)
                || string.Equals(role, AssemblyConstants.DrawerFrontWallRole, StringComparison.OrdinalIgnoreCase)
                || string.Equals(role, AssemblyConstants.DrawerBackWallRole, StringComparison.OrdinalIgnoreCase))
            {
                return new SolidColorBrush(Color.FromArgb(165, 99, 102, 241));
            }

            return UnknownFill;
        }

        private static string BuildPanelLabel(AssemblyPanel panel)
        {
            if (panel == null)
            {
                return string.Empty;
            }

            if (string.Equals(panel.PartRole, "BackPanel", StringComparison.OrdinalIgnoreCase))
            {
                return "Задняя стенка";
            }

            if (string.Equals(panel.PartRole, "VerticalPanel", StringComparison.OrdinalIgnoreCase))
            {
                return "Стойка";
            }

            if (string.Equals(panel.PartRole, "HorizontalPanel", StringComparison.OrdinalIgnoreCase))
            {
                return "Полка";
            }

            if (string.Equals(panel.PartRole, "FrontPanel", StringComparison.OrdinalIgnoreCase))
            {
                return "Передняя стенка";
            }

            if (string.Equals(panel.PartRole, AssemblyConstants.DrawerFrontRole, StringComparison.OrdinalIgnoreCase))
            {
                return "Фасад ящика";
            }

            if (string.Equals(panel.PartRole, AssemblyConstants.DrawerBottomRole, StringComparison.OrdinalIgnoreCase))
            {
                return "Дно ящика";
            }

            if (string.Equals(panel.PartRole, AssemblyConstants.DrawerSideRole, StringComparison.OrdinalIgnoreCase)
                || string.Equals(panel.PartRole, AssemblyConstants.DrawerFrontWallRole, StringComparison.OrdinalIgnoreCase)
                || string.Equals(panel.PartRole, AssemblyConstants.DrawerBackWallRole, StringComparison.OrdinalIgnoreCase))
            {
                return "Корпус ящика";
            }

            return "Панель";
        }
    }
}
