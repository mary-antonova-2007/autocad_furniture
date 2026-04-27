using System;
using Autodesk.AutoCAD.Windows;

namespace AutoCAD_BoardSorter.Ui
{
    internal static class AssemblyPalette
    {
        private static readonly Guid PaletteId = new Guid("5927E0A0-88D9-4387-A8B1-CA868E601F14");
        private static PaletteSet paletteSet;
        private static AssemblyPaletteControl control;

        public static void Show()
        {
            EnsureCreated();
            paletteSet.Visible = true;
        }

        private static void EnsureCreated()
        {
            if (paletteSet != null)
            {
                return;
            }

            control = new AssemblyPaletteControl();
            paletteSet = new PaletteSet("Сборки", PaletteId)
            {
                Style = PaletteSetStyles.ShowAutoHideButton
                    | PaletteSetStyles.ShowCloseButton
                    | PaletteSetStyles.ShowPropertiesMenu,
                MinimumSize = new System.Drawing.Size(300, 300),
                Size = new System.Drawing.Size(340, 420),
                DockEnabled = DockSides.Left | DockSides.Right
            };

            paletteSet.AddVisual("Сборки", control);
        }
    }
}
