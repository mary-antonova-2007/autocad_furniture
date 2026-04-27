using System;
using System.Collections.Generic;
using System.Windows;
using System.Windows.Media;
using AutoCAD_BoardSorter.Models;

namespace AutoCAD_BoardSorter.Ui
{
    internal interface IAssemblyEditorBackend
    {
        AssemblyScene Reload(AssemblyScene scene);
        AssemblyScene ApplyInsert(AssemblyEditorInsertRequest request);
    }

    internal sealed class AssemblyEditorInsertRequest
    {
        public AssemblyScene Scene { get; set; }
        public AssemblyNiche Niche { get; set; }
        public AssemblyEditorTool Tool { get; set; }
        public string Material { get; set; }
        public double Thickness { get; set; }
        public Rect ModelRect { get; set; }
        public IList<Rect> ModelRects { get; set; }
        public IList<AssemblyDrawerSegment> DrawerSegments { get; set; }
        public bool IsFlush { get; set; }
        public string FlushSide { get; set; }
        public double AvailableDepth { get; set; }
        public double DepthOffset { get; set; }
    }

    internal sealed class AssemblyEditorProjection
    {
        public Rect ViewportBounds { get; set; }
        public double Scale { get; set; }
        public double OffsetX { get; set; }
        public double OffsetY { get; set; }
        public double MinX { get; set; }
        public double MaxY { get; set; }

        public Point ModelToScreen(double x, double y)
        {
            return new Point(
                OffsetX + ((x - MinX) * Scale),
                OffsetY + ((MaxY - y) * Scale));
        }

        public Point ScreenToModel(Point point)
        {
            return new Point(
                MinX + ((point.X - OffsetX) / Scale),
                MaxY - ((point.Y - OffsetY) / Scale));
        }

        public Rect ModelRectToScreen(double minX, double minY, double maxX, double maxY)
        {
            Point topLeft = ModelToScreen(minX, maxY);
            Point bottomRight = ModelToScreen(maxX, minY);
            return new Rect(topLeft, bottomRight);
        }
    }

    internal sealed class AssemblyEditorRenderModel
    {
        public AssemblyEditorProjection Projection { get; set; }
        public Rect ContainerRect { get; set; }
        public readonly List<AssemblyEditorPanelVisual> Panels = new List<AssemblyEditorPanelVisual>();
        public readonly List<AssemblyEditorNicheVisual> Niches = new List<AssemblyEditorNicheVisual>();
        public AssemblyEditorPreviewVisual Preview { get; set; }
    }

    internal sealed class AssemblyEditorPanelVisual
    {
        public AssemblyPanel Source { get; set; }
        public Rect Bounds { get; set; }
        public Brush Fill { get; set; }
        public Brush Stroke { get; set; }
        public string Label { get; set; }
    }

    internal sealed class AssemblyEditorNicheVisual
    {
        public AssemblyNiche Source { get; set; }
        public Rect Bounds { get; set; }
        public Brush Fill { get; set; }
        public Brush Stroke { get; set; }
        public bool IsHovered { get; set; }
        public bool IsSelected { get; set; }
        public string Label { get; set; }
    }

    internal sealed class AssemblyEditorPreviewVisual
    {
        public Rect Bounds { get; set; }
        public Rect ModelBounds { get; set; }
        public IList<Rect> ModelRects { get; set; }
        public IList<AssemblyDrawerSegment> DrawerSegments { get; set; }
        public Brush Fill { get; set; }
        public Brush Stroke { get; set; }
        public string Label { get; set; }
        public bool IsFlush { get; set; }
        public string FlushSide { get; set; }
        public double AvailableDepth { get; set; }
        public double DepthOffset { get; set; }
    }
}
