using System.Windows;

namespace AutoCAD_BoardSorter.Ui
{
    internal sealed class AssemblyEditorHitTestService
    {
        public AssemblyEditorNicheVisual HitNiche(AssemblyEditorRenderModel render, Point point)
        {
            if (render == null)
            {
                return null;
            }

            for (int i = render.Niches.Count - 1; i >= 0; i--)
            {
                AssemblyEditorNicheVisual niche = render.Niches[i];
                if (niche.Bounds.Contains(point))
                {
                    return niche;
                }
            }

            return null;
        }
    }
}
