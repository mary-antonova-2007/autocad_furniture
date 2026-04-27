using System;
using System.Collections.Generic;
using System.Globalization;
using System.Windows;
using System.Windows.Media;
using AutoCAD_BoardSorter.Models;

namespace AutoCAD_BoardSorter.Ui
{
    internal sealed class AssemblyEditorPreviewBuilder
    {
        private static readonly Brush PreviewFill = new SolidColorBrush(Color.FromArgb(72, 244, 114, 182));
        private static readonly Brush PreviewStroke = new SolidColorBrush(Color.FromRgb(244, 114, 182));
        private static readonly Brush ErrorFill = new SolidColorBrush(Color.FromArgb(80, 239, 68, 68));
        private static readonly Brush ErrorStroke = new SolidColorBrush(Color.FromRgb(248, 113, 113));
        private const double FlushSnapModel = 24.0;
        private readonly AssemblyDrawerLayoutCalculator drawerLayoutCalculator = new AssemblyDrawerLayoutCalculator();

        public AssemblyEditorPreviewVisual Build(
            AssemblyEditorState state,
            AssemblyEditorProjection projection,
            AssemblyNiche niche,
            Point modelPoint)
        {
            if (state == null || projection == null || niche == null || state.Tool == AssemblyEditorTool.Select)
            {
                return null;
            }

            double thickness = Math.Max(1.0, state.ActiveThickness);
            Rect modelRect;
            List<Rect> modelRects = null;
            bool isFlush = false;
            string flushSide = string.Empty;
            double depthOffset = 0.0;
            AssemblyDrawerLayout drawerLayout = null;

            switch (state.Tool)
            {
                case AssemblyEditorTool.VerticalPanel:
                    modelRect = BuildVerticalPreviewRect(niche, modelPoint.X, state.ActiveOffset, Math.Min(thickness, niche.Width), out isFlush, out flushSide);
                    modelRects = BuildEqualSplitRects(state, niche, thickness, true, modelPoint);
                    break;
                case AssemblyEditorTool.HorizontalPanel:
                    modelRect = BuildHorizontalPreviewRect(niche, modelPoint.Y, state.ActiveOffset, Math.Min(thickness, niche.Height), out isFlush, out flushSide);
                    modelRects = BuildEqualSplitRects(state, niche, thickness, false, modelPoint);
                    break;
                case AssemblyEditorTool.FrontPanel:
                    modelRect = new Rect(
                        new Point(niche.MinX, niche.MinY),
                        new Point(niche.MaxX, niche.MaxY));
                    isFlush = true;
                    flushSide = "передняя плоскость";
                    depthOffset = Math.Max(0.0, state.ActiveOffset);
                    break;
                case AssemblyEditorTool.BackPanel:
                    modelRect = new Rect(
                        new Point(niche.MinX, niche.MinY),
                        new Point(niche.MaxX, niche.MaxY));
                    isFlush = true;
                    flushSide = "задняя плоскость";
                    depthOffset = Math.Max(0.0, state.ActiveOffset);
                    break;
                case AssemblyEditorTool.Drawers:
                    drawerLayout = drawerLayoutCalculator.Build(niche, state.DrawerSettings);
                    if (drawerLayout == null || drawerLayout.Bounds.IsEmpty)
                    {
                        return null;
                    }

                    modelRect = drawerLayout.Bounds;
                    modelRects = drawerLayout.FrontRects;
                    isFlush = false;
                    flushSide = "ящики";
                    break;
                default:
                    return null;
            }

            if (modelRects == null || modelRects.Count == 0)
            {
                modelRects = new List<Rect> { modelRect };
            }

            if (modelRect.IsEmpty || modelRect.Width <= 0.0 || modelRect.Height <= 0.0)
            {
                return null;
            }

            Rect screenRect = projection.ModelRectToScreen(
                modelRect.Left,
                modelRect.Top,
                modelRect.Right,
                modelRect.Bottom);

            return new AssemblyEditorPreviewVisual
            {
                Bounds = screenRect,
                ModelBounds = modelRect,
                ModelRects = modelRects,
                DrawerSegments = drawerLayout != null ? drawerLayout.Segments : null,
                Fill = drawerLayout != null && !string.IsNullOrWhiteSpace(drawerLayout.Error) ? ErrorFill : PreviewFill,
                Stroke = drawerLayout != null && !string.IsNullOrWhiteSpace(drawerLayout.Error) ? ErrorStroke : PreviewStroke,
                IsFlush = isFlush,
                FlushSide = flushSide,
                AvailableDepth = niche.Depth,
                DepthOffset = depthOffset,
                Label = drawerLayout != null
                    ? (string.IsNullOrWhiteSpace(drawerLayout.Error) ? drawerLayout.Label : drawerLayout.Error)
                    : BuildPreviewLabel(state, modelRect, modelRects.Count, niche.Depth, depthOffset, isFlush, flushSide)
            };
        }

        private static List<Rect> BuildEqualSplitRects(
            AssemblyEditorState state,
            AssemblyNiche niche,
            double thickness,
            bool isVertical,
            Point modelPoint)
        {
            if (state == null || !state.IsShiftPressed)
            {
                return null;
            }

            double totalSize = isVertical ? niche.Width : niche.Height;
            if (totalSize <= thickness + 1.0)
            {
                return null;
            }

            int partsCount = 0;
            if (state.ActiveOffset > 0.0)
            {
                partsCount = Math.Max(2, (int)Math.Round(state.ActiveOffset));
            }
            else
            {
                double cursorVal = isVertical ? modelPoint.X : modelPoint.Y;
                double distToStart = Math.Abs(cursorVal - (isVertical ? niche.MinX : niche.MinY));
                double distToEnd = Math.Abs((isVertical ? niche.MaxX : niche.MaxY) - cursorVal);
                double rawDistance = Math.Max(1.0, Math.Min(distToStart, distToEnd));
                partsCount = Math.Max(2, (int)Math.Round(totalSize / rawDistance));
            }

            double totalPanelThickness = (partsCount - 1) * thickness;
            double partSize = (totalSize - totalPanelThickness) / partsCount;
            if (partSize <= 1.0)
            {
                return null;
            }

            var result = new List<Rect>();
            double currentPos = partSize;
            for (int i = 0; i < partsCount - 1; i++)
            {
                if (isVertical)
                {
                    result.Add(NormalizeRect(currentPos + niche.MinX, niche.MinY, currentPos + niche.MinX + thickness, niche.MaxY));
                }
                else
                {
                    result.Add(NormalizeRect(niche.MinX, currentPos + niche.MinY, niche.MaxX, currentPos + niche.MinY + thickness));
                }

                currentPos += partSize + thickness;
            }

            return result;
        }

        private static Rect BuildVerticalPreviewRect(
            AssemblyNiche niche,
            double modelX,
            double requestedOffset,
            double thickness,
            out bool isFlush,
            out string flushSide)
        {
            double leftGap = Math.Abs(modelX - niche.MinX);
            double rightGap = Math.Abs(niche.MaxX - modelX);
            double minX;

            if (requestedOffset > 0.0)
            {
                bool anchorLeft = leftGap <= rightGap;
                isFlush = false;
                flushSide = anchorLeft ? "от левой стенки" : "от правой стенки";
                minX = anchorLeft ? niche.MinX + requestedOffset : niche.MaxX - requestedOffset - thickness;
            }
            else if (leftGap <= FlushSnapModel || rightGap <= FlushSnapModel)
            {
                isFlush = leftGap <= rightGap;
                flushSide = isFlush ? "левая стенка" : "правая стенка";
                minX = isFlush ? niche.MinX : (niche.MaxX - thickness);
            }
            else
            {
                isFlush = false;
                flushSide = "смещение";
                minX = modelX - (thickness * 0.5);
            }

            minX = Math.Max(niche.MinX, Math.Min(niche.MaxX - thickness, minX));
            return NormalizeRect(minX, niche.MinY, minX + thickness, niche.MaxY);
        }

        private static Rect BuildHorizontalPreviewRect(
            AssemblyNiche niche,
            double modelY,
            double requestedOffset,
            double thickness,
            out bool isFlush,
            out string flushSide)
        {
            double bottomGap = Math.Abs(modelY - niche.MinY);
            double topGap = Math.Abs(niche.MaxY - modelY);
            double minY;

            if (requestedOffset > 0.0)
            {
                bool anchorBottom = bottomGap <= topGap;
                isFlush = false;
                flushSide = anchorBottom ? "от нижней стенки" : "от верхней стенки";
                minY = anchorBottom ? niche.MinY + requestedOffset : niche.MaxY - requestedOffset - thickness;
            }
            else if (bottomGap <= FlushSnapModel || topGap <= FlushSnapModel)
            {
                isFlush = bottomGap <= topGap;
                flushSide = isFlush ? "нижняя стенка" : "верхняя стенка";
                minY = isFlush ? niche.MinY : (niche.MaxY - thickness);
            }
            else
            {
                isFlush = false;
                flushSide = "смещение";
                minY = modelY - (thickness * 0.5);
            }

            minY = Math.Max(niche.MinY, Math.Min(niche.MaxY - thickness, minY));
            return NormalizeRect(niche.MinX, minY, niche.MaxX, minY + thickness);
        }

        private static Rect NormalizeRect(double minX, double minY, double maxX, double maxY)
        {
            Point first = new Point(Math.Min(minX, maxX), Math.Min(minY, maxY));
            Point second = new Point(Math.Max(minX, maxX), Math.Max(minY, maxY));
            return new Rect(first, second);
        }

        private static string BuildPreviewLabel(
            AssemblyEditorState state,
            Rect modelRect,
            int rectCount,
            double depth,
            double depthOffset,
            bool isFlush,
            string flushSide)
        {
            string sizeText = string.Equals(flushSide, "задняя плоскость", StringComparison.OrdinalIgnoreCase)
                    || string.Equals(flushSide, "передняя плоскость", StringComparison.OrdinalIgnoreCase)
                ? string.Format(CultureInfo.InvariantCulture, "{0:0.#} x {1:0.#}", modelRect.Width, modelRect.Height)
                : string.Format(CultureInfo.InvariantCulture, "{0:0.#} мм", state.ActiveThickness);

            string depthText = depthOffset > 0.0
                ? string.Format(CultureInfo.InvariantCulture, "Глубина {0:0.#} • Отступ {1:0.#}", depth, depthOffset)
                : string.Format(CultureInfo.InvariantCulture, "Глубина {0:0.#}", depth);

            return string.Format(
                CultureInfo.InvariantCulture,
                "{0}\n{1}{2}\n{3}",
                ResolveToolTitle(state.Tool),
                isFlush ? flushSide : sizeText,
                rectCount > 1 ? string.Format(CultureInfo.InvariantCulture, "\n{0} шт.", rectCount) : string.Empty,
                depthText);
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
                    return string.Empty;
            }
        }
    }
}
