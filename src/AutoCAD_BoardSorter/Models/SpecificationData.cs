namespace AutoCAD_BoardSorter.Models
{
    internal sealed class SpecificationData
    {
        public string AssemblyNumber { get; set; }
        public string PartNumber { get; set; }
        public string PartName { get; set; }
        public string PartType { get; set; }
        public double LengthMm { get; set; }
        public double WidthMm { get; set; }
        public bool RotateLengthWidth { get; set; }
        public string Material { get; set; }
        public string Note { get; set; }
    }
}
