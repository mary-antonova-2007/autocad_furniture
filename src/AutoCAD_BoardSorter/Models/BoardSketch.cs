using System.Collections.Generic;

namespace AutoCAD_BoardSorter.Models
{
    internal sealed class BoardSketch
    {
        public readonly List<BoardSketchPoint> Points = new List<BoardSketchPoint>();
        public readonly List<BoardSketchEdge> Edges = new List<BoardSketchEdge>();

        public bool HasGeometry
        {
            get { return Points.Count >= 3 && Edges.Count >= 3; }
        }
    }

    internal sealed class BoardSketchPoint
    {
        public double X { get; set; }
        public double Y { get; set; }
    }

    internal sealed class BoardSketchEdge
    {
        public int StartIndex { get; set; }
        public int EndIndex { get; set; }
        public string Coating { get; set; }
    }
}
