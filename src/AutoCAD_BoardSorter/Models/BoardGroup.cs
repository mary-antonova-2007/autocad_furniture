using System.Collections.Generic;
using System.Linq;

namespace AutoCAD_BoardSorter.Models
{
    internal sealed class BoardGroup
    {
        public int Number { get; set; }
        public string Layer { get; set; }
        public double LengthMm { get; set; }
        public double WidthMm { get; set; }
        public double ThicknessMm { get; set; }
        public string Method { get; set; }
        public bool RotateLengthWidth { get; set; }
        public BoardCoatingSlots Coatings { get; set; }
        public BoardSketch Sketch { get; set; }
        public readonly List<BoardInfo> Items = new List<BoardInfo>();

        public int Quantity
        {
            get { return Items.Count; }
        }

        public string Handles
        {
            get { return string.Join(",", Items.Select(x => x.Handle)); }
        }
    }
}
