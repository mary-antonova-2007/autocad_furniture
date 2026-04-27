using System.Collections.Generic;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;

namespace AutoCAD_BoardSorter.Models
{
    internal enum AssemblyEditorTool
    {
        Select,
        VerticalPanel,
        HorizontalPanel,
        FrontPanel,
        BackPanel,
        Drawers
    }

    internal sealed class AssemblyScene
    {
        public AssemblyContainer Container { get; set; }
        public readonly List<AssemblyPanel> Panels = new List<AssemblyPanel>();
        public readonly List<ObjectId> IgnoredSolids = new List<ObjectId>();
        public readonly List<AssemblyNiche> Niches = new List<AssemblyNiche>();
        public readonly List<string> Warnings = new List<string>();
    }

    internal sealed class AssemblyContainer
    {
        public ObjectId ObjectId { get; set; }
        public string Handle { get; set; }
        public string AssemblyNumber { get; set; }
        public string FrontFaceKey { get; set; }
        public Vector3d FrontAxis { get; set; }
        public Vector3d UpAxis { get; set; }
        public Vector3d DepthAxis { get; set; }
        public Extents3d Bounds { get; set; }
        public Point3d Origin { get; set; }
        public double Width { get; set; }
        public double Height { get; set; }
        public double Depth { get; set; }
    }

    internal sealed class AssemblyPanel
    {
        public ObjectId ObjectId { get; set; }
        public string Handle { get; set; }
        public string PartRole { get; set; }
        public string Material { get; set; }
        public double Thickness { get; set; }
        public double Length { get; set; }
        public double Width { get; set; }
        public double MinX { get; set; }
        public double MaxX { get; set; }
        public double MinY { get; set; }
        public double MaxY { get; set; }
        public double MinDepth { get; set; }
        public double MaxDepth { get; set; }
        public bool IsRecognized { get; set; }
    }

    internal sealed class AssemblyNiche
    {
        public string Id { get; set; }
        public double MinX { get; set; }
        public double MaxX { get; set; }
        public double MinY { get; set; }
        public double MaxY { get; set; }
        public double Depth { get; set; }

        public double Width
        {
            get { return MaxX - MinX; }
        }

        public double Height
        {
            get { return MaxY - MinY; }
        }
    }
}
