using System;
using System.Globalization;
using System.IO;
using System.Text;
using Autodesk.AutoCAD.DatabaseServices;

namespace AutoCAD_BoardSorter
{
    internal static class PaletteDebugLogger
    {
        public static void Info(Database db, string message)
        {
            Write(db, "INFO ", message, null);
        }

        public static void Error(Database db, string message, Exception exception)
        {
            Write(db, "ERROR", message, exception);
        }

        private static void Write(Database db, string level, string message, Exception exception)
        {
            try
            {
                string path = GetPath(db);
                var builder = new StringBuilder();
                builder.AppendFormat(CultureInfo.InvariantCulture, "{0:O} {1} {2}", DateTime.Now, level, message);
                builder.AppendLine();
                if (exception != null)
                {
                    builder.AppendLine(exception.ToString());
                }

                File.AppendAllText(path, builder.ToString(), Encoding.UTF8);
            }
            catch
            {
                // Logging must never interfere with AutoCAD selection.
            }
        }

        private static string GetPath(Database db)
        {
            string drawingPath = db != null ? db.Filename : null;
            string directory = string.IsNullOrWhiteSpace(drawingPath)
                ? Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory)
                : Path.GetDirectoryName(drawingPath);

            string name = string.IsNullOrWhiteSpace(drawingPath)
                ? "BoardSort"
                : Path.GetFileNameWithoutExtension(drawingPath);

            string stamp = DateTime.Now.ToString("yyyyMMdd", CultureInfo.InvariantCulture);
            return Path.Combine(directory, name + "_palette_" + stamp + ".log");
        }
    }
}
