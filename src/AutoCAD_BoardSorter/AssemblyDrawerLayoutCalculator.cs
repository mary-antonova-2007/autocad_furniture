using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Windows;
using AutoCAD_BoardSorter.Models;

namespace AutoCAD_BoardSorter
{
    internal sealed class AssemblyDrawerLayoutCalculator
    {
        public const double DrawerMinClearance = 15.0;
        private const double DrawerDepthStep = 50.0;
        private const double DrawerDepthOffset = -10.0;
        private const double DrawerDepthMin = 240.0;
        private const double BodyWidthInset = 10.0;
        private const double OverlayFrontOffset = 2.0;
        private const double SideThickness = 16.0;
        private const double FrontBackThickness = 16.0;
        private const double BodyBottomOffset = 16.0;
        private const double BodyTopOffset = 10.0;
        private const double BottomBottomOffset = 13.0;
        private const double BottomSideEmbed = 7.0;
        private const double GrooveDepth = 8.0;
        private const double GrooveBottomOffset = 12.0;
        private const double GrooveExtraHeight = 1.0;

        public AssemblyDrawerLayout Build(AssemblyNiche niche, AssemblyDrawerSettings settings)
        {
            var layout = new AssemblyDrawerLayout();
            if (niche == null || settings == null || settings.Count <= 0)
            {
                layout.Error = "Ниша или настройки ящиков не заданы.";
                return layout;
            }

            EnsureDrafts(settings);

            double frontThickness = Math.Max(1.0, settings.FrontThickness);
            double signedLeft = ResolveSignedGap(settings.Mode, settings.GapLeft);
            double signedRight = ResolveSignedGap(settings.Mode, settings.GapRight);
            double signedTop = ResolveSignedGap(settings.Mode, settings.GapTop);
            double signedBottom = ResolveSignedGap(settings.Mode, settings.GapBottom);
            double gapBetween = Math.Max(0.0, settings.GapBetween);
            double frontGap = settings.Mode == AssemblyDrawerMode.Inset ? Math.Max(0.0, settings.FrontGap) : 0.0;
            double frontInset = settings.Mode == AssemblyDrawerMode.Inset ? frontGap + frontThickness : 0.0;
            double nicheDepth = Math.Max(1.0, niche.Depth - frontInset);
            double bodyDepth = settings.AutoDepth ? ComputeAutoDrawerDepth(nicheDepth) : ClampDrawerDepth(nicheDepth, settings.Depth);
            double frontWidth = Math.Max(1.0, niche.Width - signedLeft - signedRight);
            double bodyWidth = Math.Max(1.0, niche.Width - BodyWidthInset);

            double available = Math.Max(
                0.0,
                niche.Height - signedTop - signedBottom - Math.Max(0, settings.Count - 1) * gapBetween);
            List<double> heights = ComputeHeights(settings.Drawers, available);
            double totalManual = settings.Drawers.Where(x => !x.AutoHeight).Sum(x => Math.Max(0.0, x.Height));
            int autoCount = settings.Drawers.Count(x => x.AutoHeight);
            double remaining = available - totalManual;
            if (autoCount == 0 && Math.Abs(remaining) > 0.5 || autoCount > 0 && remaining < -0.5)
            {
                layout.Error = "Сумма высот ящиков больше доступной высоты.";
                SetErrorBounds(layout, niche);
                return layout;
            }

            double yCursor = niche.MaxY - signedTop;
            for (int i = 0; i < settings.Count; i++)
            {
                double frontHeight = Math.Max(1.0, heights[i]);
                double frontMaxY = yCursor;
                double frontMinY = frontMaxY - frontHeight;
                double frontMinX = niche.MinX + signedLeft;
                double frontMaxX = frontMinX + frontWidth;
                double frontMinZ = settings.Mode == AssemblyDrawerMode.Inset ? frontGap : (-frontThickness - OverlayFrontOffset);
                double frontMaxZ = settings.Mode == AssemblyDrawerMode.Inset ? (frontGap + frontThickness) : -OverlayFrontOffset;

                layout.Segments.Add(new AssemblyDrawerSegment
                {
                    DrawerIndex = i,
                    Role = AssemblyConstants.DrawerFrontRole,
                    Name = "Фасад ящика " + (i + 1).ToString(CultureInfo.InvariantCulture),
                    Material = settings.FrontMaterial,
                    MinX = frontMinX,
                    MaxX = frontMaxX,
                    MinY = frontMinY,
                    MaxY = frontMaxY,
                    MinZ = frontMinZ,
                    MaxZ = frontMaxZ
                });
                layout.FrontRects.Add(NormalizeRect(frontMinX, frontMinY, frontMaxX, frontMaxY));

                double bottomOverlayCompensation = settings.Mode == AssemblyDrawerMode.Overlay && i == settings.Count - 1 ? Math.Abs(settings.GapBottom) : 0.0;
                double topOverlayCompensation = settings.Mode == AssemblyDrawerMode.Overlay && i == 0 ? Math.Abs(settings.GapTop) : 0.0;
                double bodyMinX = niche.MinX + (BodyWidthInset * 0.5);
                double bodyMaxX = bodyMinX + bodyWidth;
                double bodyMinY = frontMinY + BodyBottomOffset + bottomOverlayCompensation;
                double bodyMaxY = frontMaxY - BodyTopOffset - topOverlayCompensation;
                double roundedBodyHeight = Math.Round(bodyMaxY - bodyMinY, 0, MidpointRounding.AwayFromZero);
                bodyMaxY = bodyMinY + Math.Max(1.0, roundedBodyHeight);
                if (bodyMaxY - bodyMinY <= 1.0)
                {
                    layout.Error = "Недостаточная высота фасада для корпуса ящика.";
                    SetErrorBounds(layout, niche);
                    return layout;
                }

                double bodyMinZ = frontInset;
                double bodyMaxZ = Math.Min(niche.Depth, bodyMinZ + bodyDepth);
                AddBodySegments(layout, i, bodyMinX, bodyMaxX, bodyMinY, bodyMaxY, bodyMinZ, bodyMaxZ, settings);

                yCursor -= frontHeight + gapBetween;
            }

            layout.Bounds = UnionRects(layout.FrontRects);
            layout.AvailableHeight = available;
            layout.Depth = bodyDepth;
            layout.Label = string.Format(
                CultureInfo.InvariantCulture,
                "Ящики\n{0} шт. • {1}\nГлубина {2:0.#}\nФасады: {3}",
                settings.Count,
                settings.Mode == AssemblyDrawerMode.Overlay ? "накладные" : "вкладные",
                bodyDepth,
                string.Join(" / ", heights.Select(x => x.ToString("0.#", CultureInfo.InvariantCulture)).ToArray()));
            return layout;
        }

        public static double ClampDrawerDepth(double nicheDepth, double desired)
        {
            double maxDepth = Math.Max(1.0, nicheDepth - DrawerMinClearance);
            if (!IsFinite(desired) || desired <= 0.0)
            {
                return maxDepth;
            }

            return Math.Max(1.0, Math.Min(desired, maxDepth));
        }

        public static double ComputeAutoDrawerDepth(double nicheDepth)
        {
            double maxDepth = Math.Max(1.0, nicheDepth - DrawerMinClearance);
            double maxStepDepth = DrawerDepthOffset + Math.Floor((maxDepth - DrawerDepthOffset) / DrawerDepthStep) * DrawerDepthStep;
            double depth = maxStepDepth >= DrawerDepthMin ? maxStepDepth : maxDepth;
            return ClampDrawerDepth(nicheDepth, depth);
        }

        public static void EnsureDrafts(AssemblyDrawerSettings settings)
        {
            if (settings == null)
            {
                return;
            }

            settings.Count = Math.Max(1, settings.Count);
            while (settings.Drawers.Count < settings.Count)
            {
                settings.Drawers.Add(new AssemblyDrawerDraft { AutoHeight = true, Height = 0.0 });
            }

            while (settings.Drawers.Count > settings.Count)
            {
                settings.Drawers.RemoveAt(settings.Drawers.Count - 1);
            }
        }

        private static void AddBodySegments(
            AssemblyDrawerLayout layout,
            int index,
            double bodyMinX,
            double bodyMaxX,
            double bodyMinY,
            double bodyMaxY,
            double bodyMinZ,
            double bodyMaxZ,
            AssemblyDrawerSettings settings)
        {
            double bodyWidth = Math.Max(1.0, bodyMaxX - bodyMinX);
            double bodyHeight = Math.Max(1.0, bodyMaxY - bodyMinY);
            double bodyDepth = Math.Max(1.0, bodyMaxZ - bodyMinZ);
            double side = Math.Min(SideThickness, Math.Max(1.0, bodyWidth * 0.5));
            double wall = Math.Min(FrontBackThickness, Math.Max(1.0, bodyDepth));
            double bottom = Math.Min(Math.Max(1.0, settings.BottomThickness), Math.Max(1.0, bodyHeight));
            double bottomOffset = Math.Min(BottomBottomOffset, Math.Max(0.0, bodyHeight - bottom));
            double bottomWidth = Math.Max(1.0, bodyWidth - (side * 2.0) + (BottomSideEmbed * 2.0));
            double bottomMinX = bodyMinX + side - BottomSideEmbed;
            double bottomMaxX = bottomMinX + bottomWidth;
            double bottomMinY = bodyMinY + bottomOffset;
            double bottomMaxY = bottomMinY + bottom;
            double wallMinY = Math.Min(bodyMaxY - 1.0, bottomMaxY);
            double grooveBottom = Math.Min(GrooveBottomOffset, Math.Max(0.0, bodyHeight));
            double grooveHeight = Math.Min(bottom + GrooveExtraHeight, Math.Max(0.0, bodyHeight - grooveBottom));
            double grooveTop = bodyMinY + grooveBottom + grooveHeight;
            double grooveThickness = Math.Max(1.0, side - Math.Min(GrooveDepth, Math.Max(0.0, side - 1.0)));

            AddSideWithGroove(layout, index, settings.BodyMaterial, bodyMinX, bodyMinX + side, bodyMinY, bodyMaxY, bodyMinZ, bodyMaxZ, grooveThickness, bodyMinY + grooveBottom, grooveTop, true);
            AddSideWithGroove(layout, index, settings.BodyMaterial, bodyMaxX - side, bodyMaxX, bodyMinY, bodyMaxY, bodyMinZ, bodyMaxZ, grooveThickness, bodyMinY + grooveBottom, grooveTop, false);
            Add(layout, index, AssemblyConstants.DrawerFrontWallRole, "Передняя стенка ящика " + (index + 1), settings.BodyMaterial, bodyMinX + side, bodyMaxX - side, wallMinY, bodyMaxY, bodyMinZ, bodyMinZ + wall);
            Add(layout, index, AssemblyConstants.DrawerBackWallRole, "Задняя стенка ящика " + (index + 1), settings.BodyMaterial, bodyMinX + side, bodyMaxX - side, wallMinY, bodyMaxY, bodyMaxZ - wall, bodyMaxZ);
            Add(layout, index, AssemblyConstants.DrawerBottomRole, "Дно ящика " + (index + 1), settings.BottomMaterial, bottomMinX, bottomMaxX, bodyMinY + bottomOffset, bodyMinY + bottomOffset + bottom, bodyMinZ, bodyMaxZ);
        }

        private static void AddSideWithGroove(
            AssemblyDrawerLayout layout,
            int index,
            string material,
            double minX,
            double maxX,
            double minY,
            double maxY,
            double minZ,
            double maxZ,
            double grooveThickness,
            double grooveMinY,
            double grooveMaxY,
            bool leftSide)
        {
            string name = "Боковина ящика " + (index + 1);
            Add(layout, index, AssemblyConstants.DrawerSideRole, name, material, minX, maxX, minY, maxY, minZ, maxZ);
            AssemblyDrawerSegment segment = layout.Segments[layout.Segments.Count - 1];
            segment.HasSideGroove = true;
            segment.GrooveMinY = grooveMinY;
            segment.GrooveMaxY = grooveMaxY;
            segment.GrooveThickness = grooveThickness;
            segment.GrooveOnMaxX = leftSide;
        }

        private static void Add(AssemblyDrawerLayout layout, int drawerIndex, string role, string name, string material, double minX, double maxX, double minY, double maxY, double minZ, double maxZ)
        {
            if (maxX - minX <= 0.1 || maxY - minY <= 0.1 || maxZ - minZ <= 0.1)
            {
                return;
            }

            layout.Segments.Add(new AssemblyDrawerSegment
            {
                DrawerIndex = drawerIndex,
                Role = role,
                Name = name,
                Material = material,
                MinX = minX,
                MaxX = maxX,
                MinY = minY,
                MaxY = maxY,
                MinZ = minZ,
                MaxZ = maxZ
            });
        }

        private static List<double> ComputeHeights(IList<AssemblyDrawerDraft> drawers, double available)
        {
            double manual = drawers.Where(x => !x.AutoHeight).Sum(x => Math.Max(0.0, x.Height));
            List<int> autoIndexes = drawers
                .Select((drawer, index) => drawer.AutoHeight ? index : -1)
                .Where(index => index >= 0)
                .ToList();
            var heights = drawers.Select(x => x.AutoHeight ? 0.0 : Math.Max(0.0, x.Height)).ToList();
            if (autoIndexes.Count == 0)
            {
                return heights;
            }

            double baseHeight = (available - manual) / autoIndexes.Count;
            foreach (int index in autoIndexes)
            {
                heights[index] = baseHeight;
            }

            int last = autoIndexes[autoIndexes.Count - 1];
            heights[last] += (available - manual) - (baseHeight * autoIndexes.Count);
            return heights;
        }

        private static double ResolveSignedGap(AssemblyDrawerMode mode, double value)
        {
            double safe = Math.Abs(IsFinite(value) ? value : 0.0);
            return mode == AssemblyDrawerMode.Overlay ? -safe : safe;
        }

        private static Rect NormalizeRect(double minX, double minY, double maxX, double maxY)
        {
            return new Rect(new Point(Math.Min(minX, maxX), Math.Min(minY, maxY)), new Point(Math.Max(minX, maxX), Math.Max(minY, maxY)));
        }

        private static Rect UnionRects(IList<Rect> rects)
        {
            if (rects == null || rects.Count == 0)
            {
                return Rect.Empty;
            }

            Rect result = rects[0];
            for (int i = 1; i < rects.Count; i++)
            {
                result.Union(rects[i]);
            }

            return result;
        }

        private static bool IsFinite(double value)
        {
            return !double.IsNaN(value) && !double.IsInfinity(value);
        }

        private static void SetErrorBounds(AssemblyDrawerLayout layout, AssemblyNiche niche)
        {
            Rect bounds = NormalizeRect(niche.MinX, niche.MinY, niche.MaxX, niche.MaxY);
            layout.Bounds = bounds;
            layout.FrontRects.Clear();
            layout.FrontRects.Add(bounds);
        }
    }

    internal sealed class AssemblyDrawerLayout
    {
        public readonly List<AssemblyDrawerSegment> Segments = new List<AssemblyDrawerSegment>();
        public readonly List<Rect> FrontRects = new List<Rect>();
        public Rect Bounds { get; set; }
        public string Label { get; set; }
        public string Error { get; set; }
        public double AvailableHeight { get; set; }
        public double Depth { get; set; }
    }

    internal sealed class AssemblyDrawerSegment
    {
        public int DrawerIndex { get; set; }
        public string Role { get; set; }
        public string Name { get; set; }
        public string Material { get; set; }
        public double MinX { get; set; }
        public double MaxX { get; set; }
        public double MinY { get; set; }
        public double MaxY { get; set; }
        public double MinZ { get; set; }
        public double MaxZ { get; set; }
        public bool HasSideGroove { get; set; }
        public bool GrooveOnMaxX { get; set; }
        public double GrooveMinY { get; set; }
        public double GrooveMaxY { get; set; }
        public double GrooveThickness { get; set; }
    }
}
