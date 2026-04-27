using Autodesk.AutoCAD.DatabaseServices;
using AutoCAD_BoardSorter.Models;

namespace AutoCAD_BoardSorter
{
    internal static class AssemblyPartStorage
    {
        private const string RecordName = "AssemblyPart";
        private const int Version = 1;
        private const int MinimumValueCount = 6;

        public static void Write(Entity entity, AssemblyPartData data, Transaction tr)
        {
            AssemblyMetadataStorage.Write(entity, tr, RecordName, ToResultBuffer(data));
        }

        public static bool TryRead(Entity entity, Transaction tr, out AssemblyPartData data)
        {
            data = null;

            ResultBuffer buffer;
            if (!AssemblyMetadataStorage.TryRead(entity, tr, RecordName, out buffer))
            {
                return false;
            }

            data = FromResultBuffer(buffer);
            return data != null;
        }

        public static bool Clear(Entity entity, Transaction tr)
        {
            return AssemblyMetadataStorage.Clear(entity, tr, RecordName);
        }

        private static ResultBuffer ToResultBuffer(AssemblyPartData data)
        {
            return AssemblyMetadataStorage.CreateTextBuffer(
                Version,
                data == null ? null : data.AssemblyNumber,
                data == null ? null : data.PartRole,
                data == null ? null : data.SourceContainerHandle,
                data == null ? null : data.GeneratedByConstructor,
                data == null ? null : data.Material);
        }

        private static AssemblyPartData FromResultBuffer(ResultBuffer buffer)
        {
            if (buffer == null)
            {
                return null;
            }

            TypedValue[] values = buffer.AsArray();
            if (values.Length < MinimumValueCount)
            {
                return null;
            }

            return new AssemblyPartData
            {
                AssemblyNumber = AssemblyMetadataStorage.ReadText(values, 1),
                PartRole = AssemblyMetadataStorage.ReadText(values, 2),
                SourceContainerHandle = AssemblyMetadataStorage.ReadText(values, 3),
                GeneratedByConstructor = AssemblyMetadataStorage.ReadText(values, 4),
                Material = AssemblyMetadataStorage.ReadText(values, 5)
            };
        }
    }
}
