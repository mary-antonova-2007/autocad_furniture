using System;
using System.Globalization;
using System.IO;
using System.Text;
using Autodesk.AutoCAD.DatabaseServices;

namespace AutoCAD_BoardSorter
{
    internal sealed class BoardSortLogger : IDisposable
    {
        private readonly StreamWriter writer;

        private BoardSortLogger(string path)
        {
            Path = path;
            writer = new StreamWriter(path, false, Encoding.UTF8);
            Info("Log started");
        }

        public string Path { get; private set; }

        public static BoardSortLogger Create(Database db, string suffix)
        {
            string drawingPath = db.Filename;
            string directory = string.IsNullOrWhiteSpace(drawingPath)
                ? Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory)
                : System.IO.Path.GetDirectoryName(drawingPath);

            string name = string.IsNullOrWhiteSpace(drawingPath)
                ? "BoardSort"
                : System.IO.Path.GetFileNameWithoutExtension(drawingPath);

            string stamp = DateTime.Now.ToString("yyyyMMdd_HHmmss", CultureInfo.InvariantCulture);
            string path = System.IO.Path.Combine(directory, name + "_" + suffix + "_" + stamp + ".log");
            return new BoardSortLogger(path);
        }

        public void Info(string message)
        {
            writer.WriteLine("{0:O} INFO  {1}", DateTime.Now, message);
            writer.Flush();
        }

        public void Error(string message, Exception exception)
        {
            writer.WriteLine("{0:O} ERROR {1}", DateTime.Now, message);
            writer.WriteLine(exception);
            writer.Flush();
        }

        public void Dispose()
        {
            Info("Log finished");
            writer.Dispose();
        }
    }
}
