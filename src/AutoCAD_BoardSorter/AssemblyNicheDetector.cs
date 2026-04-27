using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using AutoCAD_BoardSorter.Models;

namespace AutoCAD_BoardSorter
{
    internal sealed class AssemblyNicheDetector
    {
        private const double Tolerance = 0.5;
        private const double MinimumNicheSize = 30.0;

        public List<AssemblyNiche> Detect(AssemblyContainer container, IList<AssemblyPanel> panels)
        {
            var result = new List<AssemblyNiche>();
            if (container == null)
            {
                return result;
            }

            List<double> xs = CollectCoordinates(0.0, container.Width, panels.SelectMany(GetXCoordinates));
            List<double> ys = CollectCoordinates(0.0, container.Height, panels.SelectMany(GetYCoordinates));
            if (xs.Count < 2 || ys.Count < 2)
            {
                return result;
            }

            int cols = xs.Count - 1;
            int rows = ys.Count - 1;
            var free = new bool[cols, rows];
            var depths = new double[cols, rows];

            for (int x = 0; x < cols; x++)
            {
                for (int y = 0; y < rows; y++)
                {
                    double centerX = (xs[x] + xs[x + 1]) * 0.5;
                    double centerY = (ys[y] + ys[y + 1]) * 0.5;
                    if (IsOccupied(centerX, centerY, panels))
                    {
                        free[x, y] = false;
                        continue;
                    }

                    free[x, y] = true;
                    depths[x, y] = ComputeDepth(container, panels, centerX, centerY);
                }
            }

            var visited = new bool[cols, rows];
            int nicheIndex = 1;
            for (int y = rows - 1; y >= 0; y--)
            {
                for (int x = 0; x < cols; x++)
                {
                    if (!free[x, y] || visited[x, y])
                    {
                        continue;
                    }

                    double depth = depths[x, y];
                    int right = x;
                    while (right + 1 < cols && free[right + 1, y] && !visited[right + 1, y] && SameDepth(depths[right + 1, y], depth))
                    {
                        right++;
                    }

                    int bottom = y;
                    bool canGrow = true;
                    while (canGrow && bottom - 1 >= 0)
                    {
                        for (int testX = x; testX <= right; testX++)
                        {
                            if (!free[testX, bottom - 1] || visited[testX, bottom - 1] || !SameDepth(depths[testX, bottom - 1], depth))
                            {
                                canGrow = false;
                                break;
                            }
                        }

                        if (canGrow)
                        {
                            bottom--;
                        }
                    }

                    for (int markX = x; markX <= right; markX++)
                    {
                        for (int markY = bottom; markY <= y; markY++)
                        {
                            visited[markX, markY] = true;
                        }
                    }

                    double minX = xs[x];
                    double maxX = xs[right + 1];
                    double minY = ys[bottom];
                    double maxY = ys[y + 1];
                    if (maxX - minX <= MinimumNicheSize || maxY - minY <= MinimumNicheSize || depth <= Tolerance)
                    {
                        continue;
                    }

                    result.Add(new AssemblyNiche
                    {
                        Id = "N" + nicheIndex.ToString(CultureInfo.InvariantCulture),
                        MinX = minX,
                        MaxX = maxX,
                        MinY = minY,
                        MaxY = maxY,
                        Depth = depth
                    });
                    nicheIndex++;
                }
            }

            return result
                .OrderByDescending(x => x.MaxY)
                .ThenBy(x => x.MinX)
                .ToList();
        }

        private static IEnumerable<double> GetXCoordinates(AssemblyPanel panel)
        {
            yield return panel.MinX;
            yield return panel.MaxX;
        }

        private static IEnumerable<double> GetYCoordinates(AssemblyPanel panel)
        {
            yield return panel.MinY;
            yield return panel.MaxY;
        }

        private static List<double> CollectCoordinates(double min, double max, IEnumerable<double> values)
        {
            var coords = new List<double> { min, max };
            foreach (double value in values)
            {
                if (value > min + Tolerance && value < max - Tolerance)
                {
                    coords.Add(value);
                }
            }

            coords.Sort();
            var result = new List<double>();
            foreach (double value in coords)
            {
                if (result.Count == 0 || Math.Abs(result[result.Count - 1] - value) > Tolerance)
                {
                    result.Add(value);
                }
            }

            return result;
        }

        private static bool IsOccupied(double x, double y, IEnumerable<AssemblyPanel> panels)
        {
            foreach (AssemblyPanel panel in panels)
            {
                if (!panel.IsRecognized
                    || string.Equals(panel.PartRole, AssemblyConstants.BackPanelRole, StringComparison.OrdinalIgnoreCase)
                    || string.Equals(panel.PartRole, AssemblyConstants.FrontPanelRole, StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                if (x > panel.MinX + Tolerance
                    && x < panel.MaxX - Tolerance
                    && y > panel.MinY + Tolerance
                    && y < panel.MaxY - Tolerance)
                {
                    return true;
                }
            }

            return false;
        }

        private static double ComputeDepth(AssemblyContainer container, IEnumerable<AssemblyPanel> panels, double x, double y)
        {
            double depth = container.Depth;
            foreach (AssemblyPanel panel in panels)
            {
                if (!panel.IsRecognized || !string.Equals(panel.PartRole, AssemblyConstants.BackPanelRole, StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                if (x >= panel.MinX - Tolerance
                    && x <= panel.MaxX + Tolerance
                    && y >= panel.MinY - Tolerance
                    && y <= panel.MaxY + Tolerance)
                {
                    depth = Math.Min(depth, panel.MinDepth);
                }
            }

            return Math.Max(0.0, depth);
        }

        private static bool SameDepth(double first, double second)
        {
            return Math.Abs(first - second) <= Tolerance;
        }
    }
}
