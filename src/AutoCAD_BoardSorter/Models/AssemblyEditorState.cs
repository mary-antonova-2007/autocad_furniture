namespace AutoCAD_BoardSorter.Models
{
    internal sealed class AssemblyEditorState
    {
        public AssemblyScene Scene { get; set; }
        public AssemblyEditorTool Tool { get; set; }
        public string ActiveMaterial { get; set; }
        public double ActiveThickness { get; set; }
        public double ActiveOffset { get; set; }
        public AssemblyNiche HoveredNiche { get; set; }
        public AssemblyNiche SelectedNiche { get; set; }
        public string StatusText { get; set; }
        public bool IsShiftPressed { get; set; }
        public double ViewZoom { get; set; }
        public double ViewPanX { get; set; }
        public double ViewPanY { get; set; }
        public AssemblyDrawerSettings DrawerSettings { get; set; }
    }
}
