namespace AutoCAD_BoardSorter.Models
{
    internal sealed class AssemblyPartData
    {
        public string AssemblyNumber { get; set; }
        public string PartRole { get; set; }
        public string SourceContainerHandle { get; set; }
        public string GeneratedByConstructor { get; set; }
        public string Material { get; set; }
    }
}
