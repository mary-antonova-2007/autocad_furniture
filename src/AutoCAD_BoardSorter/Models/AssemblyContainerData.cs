namespace AutoCAD_BoardSorter.Models
{
    internal sealed class AssemblyContainerData
    {
        public string AssemblyNumber { get; set; }
        public string EntityRole { get; set; }
        public string Version { get; set; }
        public string FrontFaceKey { get; set; }
        public string FrontAxisX { get; set; }
        public string FrontAxisY { get; set; }
        public string FrontAxisZ { get; set; }
        public string UpAxisX { get; set; }
        public string UpAxisY { get; set; }
        public string UpAxisZ { get; set; }
        public string DepthAxisX { get; set; }
        public string DepthAxisY { get; set; }
        public string DepthAxisZ { get; set; }
        public string Width { get; set; }
        public string Height { get; set; }
        public string Depth { get; set; }
    }
}
