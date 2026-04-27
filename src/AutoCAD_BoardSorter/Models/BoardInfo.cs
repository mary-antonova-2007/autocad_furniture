namespace AutoCAD_BoardSorter.Models
{
    using Autodesk.AutoCAD.DatabaseServices;
    using System.Globalization;

    internal sealed class BoardInfo
    {
        public ObjectId ObjectId { get; set; }
        public int Number { get; set; }
        public string Handle { get; set; }
        public string Layer { get; set; }
        public double LengthMm { get; set; }
        public double WidthMm { get; set; }
        public double ThicknessMm { get; set; }
        public string Method { get; set; }
        public string Fingerprint { get; set; }
        public bool RotateLengthWidth { get; set; }
        public string AssemblyNumber { get; set; }
        public string PartName { get; set; }
        public string Material { get; set; }
        public BoardCoatingSlots Coatings { get; set; }
        public BoardSketch Sketch { get; set; }

        public string ThicknessKey
        {
            get { return ThicknessMm.ToString("0.0", CultureInfo.InvariantCulture); }
        }
    }
}
