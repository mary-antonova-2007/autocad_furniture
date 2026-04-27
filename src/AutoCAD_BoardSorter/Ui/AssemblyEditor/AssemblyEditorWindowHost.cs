using System.Windows;
using Autodesk.AutoCAD.DatabaseServices;
using AutoCAD_BoardSorter.Models;

namespace AutoCAD_BoardSorter.Ui
{
    internal static class AssemblyEditorWindowHost
    {
        private static AssemblyEditorWindow window;
        private static ObjectId currentContainerId = ObjectId.Null;

        public static void Show(AssemblyScene scene = null, IAssemblyEditorBackend backend = null)
        {
            if (window == null || !window.IsLoaded)
            {
                window = new AssemblyEditorWindow(backend);
                window.Closed += delegate
                {
                    window = null;
                    currentContainerId = ObjectId.Null;
                };
            }

            window.LoadScene(scene ?? AssemblyEditorSampleSceneFactory.Create());
            if (!window.IsVisible)
            {
                window.Show();
            }

            if (window.WindowState == WindowState.Minimized)
            {
                window.WindowState = WindowState.Normal;
            }

            window.Activate();
        }

        public static void Show(ObjectId containerId)
        {
            if (containerId.IsNull)
            {
                return;
            }

            if (window != null && window.IsLoaded && currentContainerId != containerId)
            {
                window.Close();
                window = null;
            }

            currentContainerId = containerId;
            var backend = new AutoCadAssemblyEditorBackend(containerId);
            AssemblyScene scene = backend.Reload(null) ?? new AssemblyScene();
            Show(scene, backend);
        }
    }
}
