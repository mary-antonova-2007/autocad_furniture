using System;
using Autodesk.AutoCAD.DatabaseServices;
using AutoCAD_BoardSorter.Models;

namespace AutoCAD_BoardSorter
{
    internal static class SpecificationStorage
    {
        private const string DictionaryName = "AutoCAD_BoardSorter";
        private const string RecordName = "Specification";
        private const int Version = 1;

        public static void Write(Entity entity, SpecificationData data, Transaction tr)
        {
            if (entity.ExtensionDictionary.IsNull)
            {
                entity.UpgradeOpen();
                entity.CreateExtensionDictionary();
            }

            var extensionDictionary = (DBDictionary)tr.GetObject(entity.ExtensionDictionary, OpenMode.ForWrite);
            DBDictionary appDictionary;

            if (extensionDictionary.Contains(DictionaryName))
            {
                appDictionary = (DBDictionary)tr.GetObject(extensionDictionary.GetAt(DictionaryName), OpenMode.ForWrite);
            }
            else
            {
                appDictionary = new DBDictionary();
                extensionDictionary.SetAt(DictionaryName, appDictionary);
                tr.AddNewlyCreatedDBObject(appDictionary, true);
            }

            if (appDictionary.Contains(RecordName))
            {
                DBObject old = tr.GetObject(appDictionary.GetAt(RecordName), OpenMode.ForWrite);
                old.Erase();
            }

            var record = new Xrecord
            {
                Data = ToResultBuffer(data)
            };

            appDictionary.SetAt(RecordName, record);
            tr.AddNewlyCreatedDBObject(record, true);
        }

        public static bool TryRead(Entity entity, Transaction tr, out SpecificationData data)
        {
            data = null;

            if (entity.ExtensionDictionary.IsNull)
            {
                return false;
            }

            var extensionDictionary = (DBDictionary)tr.GetObject(entity.ExtensionDictionary, OpenMode.ForRead);
            if (!extensionDictionary.Contains(DictionaryName))
            {
                return false;
            }

            var appDictionary = (DBDictionary)tr.GetObject(extensionDictionary.GetAt(DictionaryName), OpenMode.ForRead);
            if (!appDictionary.Contains(RecordName))
            {
                return false;
            }

            var record = (Xrecord)tr.GetObject(appDictionary.GetAt(RecordName), OpenMode.ForRead);
            data = FromResultBuffer(record.Data);
            return data != null;
        }

        public static bool Clear(Entity entity, Transaction tr)
        {
            if (entity.ExtensionDictionary.IsNull)
            {
                return false;
            }

            var extensionDictionary = (DBDictionary)tr.GetObject(entity.ExtensionDictionary, OpenMode.ForRead);
            if (!extensionDictionary.Contains(DictionaryName))
            {
                return false;
            }

            var appDictionary = (DBDictionary)tr.GetObject(extensionDictionary.GetAt(DictionaryName), OpenMode.ForWrite);
            if (!appDictionary.Contains(RecordName))
            {
                return false;
            }

            DBObject record = tr.GetObject(appDictionary.GetAt(RecordName), OpenMode.ForWrite);
            record.Erase();
            return true;
        }

        private static ResultBuffer ToResultBuffer(SpecificationData data)
        {
            return new ResultBuffer(
                new TypedValue((int)DxfCode.Int16, Version),
                new TypedValue((int)DxfCode.Text, data.AssemblyNumber ?? string.Empty),
                new TypedValue((int)DxfCode.Text, data.PartNumber ?? string.Empty),
                new TypedValue((int)DxfCode.Text, data.PartName ?? string.Empty),
                new TypedValue((int)DxfCode.Text, data.PartType ?? string.Empty),
                new TypedValue((int)DxfCode.Real, data.LengthMm),
                new TypedValue((int)DxfCode.Real, data.WidthMm),
                new TypedValue((int)DxfCode.Int16, data.RotateLengthWidth ? 1 : 0),
                new TypedValue((int)DxfCode.Text, data.Material ?? string.Empty),
                new TypedValue((int)DxfCode.Text, data.Note ?? string.Empty));
        }

        private static SpecificationData FromResultBuffer(ResultBuffer buffer)
        {
            if (buffer == null)
            {
                return null;
            }

            TypedValue[] values = buffer.AsArray();
            if (values.Length < 10)
            {
                return null;
            }

            return new SpecificationData
            {
                AssemblyNumber = Convert.ToString(values[1].Value),
                PartNumber = Convert.ToString(values[2].Value),
                PartName = Convert.ToString(values[3].Value),
                PartType = Convert.ToString(values[4].Value),
                LengthMm = Convert.ToDouble(values[5].Value),
                WidthMm = Convert.ToDouble(values[6].Value),
                RotateLengthWidth = Convert.ToInt16(values[7].Value) != 0,
                Material = Convert.ToString(values[8].Value),
                Note = Convert.ToString(values[9].Value)
            };
        }
    }
}
