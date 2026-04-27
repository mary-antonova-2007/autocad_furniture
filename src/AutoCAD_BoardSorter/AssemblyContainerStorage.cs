using Autodesk.AutoCAD.DatabaseServices;
using AutoCAD_BoardSorter.Models;

namespace AutoCAD_BoardSorter
{
    internal static class AssemblyContainerStorage
    {
        private const string RecordName = "AssemblyContainer";
        private const int Version = 1;
        private const int MinimumValueCount = 14;

        public static void Write(Entity entity, AssemblyContainerData data, Transaction tr)
        {
            AssemblyMetadataStorage.Write(entity, tr, RecordName, ToResultBuffer(data));
        }

        public static bool TryRead(Entity entity, Transaction tr, out AssemblyContainerData data)
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

        private static ResultBuffer ToResultBuffer(AssemblyContainerData data)
        {
            return AssemblyMetadataStorage.CreateTextBuffer(
                Version,
                data == null ? null : data.AssemblyNumber,
                data == null ? null : data.EntityRole,
                data == null ? null : data.Version,
                data == null ? null : data.FrontFaceKey,
                data == null ? null : data.FrontAxisX,
                data == null ? null : data.FrontAxisY,
                data == null ? null : data.FrontAxisZ,
                data == null ? null : data.UpAxisX,
                data == null ? null : data.UpAxisY,
                data == null ? null : data.UpAxisZ,
                data == null ? null : data.DepthAxisX,
                data == null ? null : data.DepthAxisY,
                data == null ? null : data.DepthAxisZ,
                data == null ? null : data.Width,
                data == null ? null : data.Height,
                data == null ? null : data.Depth);
        }

        private static AssemblyContainerData FromResultBuffer(ResultBuffer buffer)
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

            var data = new AssemblyContainerData
            {
                AssemblyNumber = AssemblyMetadataStorage.ReadText(values, 1),
                EntityRole = AssemblyMetadataStorage.ReadText(values, 2),
                Version = AssemblyMetadataStorage.ReadText(values, 3),
                FrontFaceKey = AssemblyMetadataStorage.ReadText(values, 4),
                FrontAxisX = AssemblyMetadataStorage.ReadText(values, 5),
                FrontAxisY = AssemblyMetadataStorage.ReadText(values, 6),
                FrontAxisZ = AssemblyMetadataStorage.ReadText(values, 7),
                UpAxisX = AssemblyMetadataStorage.ReadText(values, 8),
                UpAxisY = AssemblyMetadataStorage.ReadText(values, 9),
                UpAxisZ = AssemblyMetadataStorage.ReadText(values, 10),
                DepthAxisX = AssemblyMetadataStorage.ReadText(values, 11),
                DepthAxisY = AssemblyMetadataStorage.ReadText(values, 12),
                DepthAxisZ = AssemblyMetadataStorage.ReadText(values, 13)
            };
            if (values.Length > 14)
            {
                data.Width = AssemblyMetadataStorage.ReadText(values, 14);
            }

            if (values.Length > 15)
            {
                data.Height = AssemblyMetadataStorage.ReadText(values, 15);
            }

            if (values.Length > 16)
            {
                data.Depth = AssemblyMetadataStorage.ReadText(values, 16);
            }

            return data;
        }
    }
}
