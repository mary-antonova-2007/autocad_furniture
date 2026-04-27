using System.Collections.Generic;

namespace AutoCAD_BoardSorter.Models
{
    internal enum AssemblyDrawerMode
    {
        Overlay,
        Inset
    }

    internal sealed class AssemblyDrawerSettings
    {
        public int Count { get; set; }
        public AssemblyDrawerMode Mode { get; set; }
        public double GapLeft { get; set; }
        public double GapRight { get; set; }
        public double GapTop { get; set; }
        public double GapBottom { get; set; }
        public double GapBetween { get; set; }
        public double FrontGap { get; set; }
        public double Depth { get; set; }
        public bool AutoDepth { get; set; }
        public string FrontMaterial { get; set; }
        public string BodyMaterial { get; set; }
        public string BottomMaterial { get; set; }
        public double FrontThickness { get; set; }
        public double BottomThickness { get; set; }
        public readonly List<AssemblyDrawerDraft> Drawers = new List<AssemblyDrawerDraft>();
    }

    internal sealed class AssemblyDrawerDraft
    {
        public double Height { get; set; }
        public bool AutoHeight { get; set; }
    }
}
