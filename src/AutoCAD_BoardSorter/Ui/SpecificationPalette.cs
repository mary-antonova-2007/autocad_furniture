using System;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Windows;

namespace AutoCAD_BoardSorter.Ui
{
    internal static class SpecificationPalette
    {
        private static readonly Guid PaletteId = new Guid("6E12597E-66BB-4F8C-A1B2-6D1E9430634D");
        private static PaletteSet paletteSet;
        private static SpecificationPaletteControl control;
        private static bool refreshQueued;
        private static bool refreshInProgress;

        public static void Show()
        {
            EnsureCreated();
            paletteSet.Visible = true;
            control.RefreshFromSelection();
        }

        private static void EnsureCreated()
        {
            if (paletteSet != null)
            {
                return;
            }

            control = new SpecificationPaletteControl();
            paletteSet = new PaletteSet("Спецификация", PaletteId)
            {
                Style = PaletteSetStyles.ShowAutoHideButton
                    | PaletteSetStyles.ShowCloseButton
                    | PaletteSetStyles.ShowPropertiesMenu,
                MinimumSize = new System.Drawing.Size(330, 420),
                Size = new System.Drawing.Size(390, 560),
                DockEnabled = DockSides.Left | DockSides.Right
            };

            paletteSet.AddVisual("Спецификация", control);
            HookEditorEvents();
        }

        private static void HookEditorEvents()
        {
            Application.DocumentManager.DocumentActivated += delegate { HookActiveEditor(); };
            HookActiveEditor();
        }

        private static void HookActiveEditor()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            if (doc == null)
            {
                return;
            }

            Editor editor = doc.Editor;
            editor.SelectionAdded -= OnSelectionChanged;
            editor.SelectionRemoved -= OnSelectionChanged;
            editor.SelectionAdded += OnSelectionChanged;
            editor.SelectionRemoved += OnSelectionChanged;
        }

        private static void OnSelectionChanged(object sender, EventArgs e)
        {
            if (paletteSet == null || !paletteSet.Visible || control == null)
            {
                return;
            }

            if (refreshInProgress || refreshQueued)
            {
                return;
            }

            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc != null ? doc.Database : null;
            PaletteDebugLogger.Info(db, "Selection event: " + e.GetType().FullName);

            refreshQueued = true;
            control.Dispatcher.BeginInvoke(new Action(delegate
            {
                refreshQueued = false;
                refreshInProgress = true;
                try
                {
                    PaletteDebugLogger.Info(db, "Deferred RefreshFromSelection start");
                    control.RefreshFromSelection();
                    PaletteDebugLogger.Info(db, "Deferred RefreshFromSelection finish");
                }
                catch (Exception ex)
                {
                    PaletteDebugLogger.Error(db, "Deferred RefreshFromSelection failed", ex);
                }
                finally
                {
                    refreshInProgress = false;
                }
            }));
        }
    }
}
