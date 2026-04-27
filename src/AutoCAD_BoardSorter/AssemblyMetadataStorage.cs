using System.Collections.Generic;
using Autodesk.AutoCAD.DatabaseServices;

namespace AutoCAD_BoardSorter
{
    internal static class AssemblyMetadataStorage
    {
        private const string DictionaryName = "AutoCAD_BoardSorter";

        public static void Write(Entity entity, Transaction tr, string recordName, ResultBuffer data)
        {
            var appDictionary = GetOrCreateAppDictionary(entity, tr);

            if (appDictionary.Contains(recordName))
            {
                DBObject old = tr.GetObject(appDictionary.GetAt(recordName), OpenMode.ForWrite);
                old.Erase();
            }

            var record = new Xrecord
            {
                Data = data
            };

            appDictionary.SetAt(recordName, record);
            tr.AddNewlyCreatedDBObject(record, true);
        }

        public static bool TryRead(Entity entity, Transaction tr, string recordName, out ResultBuffer data)
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
            if (!appDictionary.Contains(recordName))
            {
                return false;
            }

            var record = (Xrecord)tr.GetObject(appDictionary.GetAt(recordName), OpenMode.ForRead);
            data = record.Data;
            return data != null;
        }

        public static bool Clear(Entity entity, Transaction tr, string recordName)
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
            if (!appDictionary.Contains(recordName))
            {
                return false;
            }

            DBObject record = tr.GetObject(appDictionary.GetAt(recordName), OpenMode.ForWrite);
            record.Erase();
            return true;
        }

        public static ResultBuffer CreateTextBuffer(int version, params string[] values)
        {
            var typedValues = new List<TypedValue>
            {
                new TypedValue((int)DxfCode.Int16, version)
            };

            foreach (string value in values)
            {
                typedValues.Add(new TypedValue((int)DxfCode.Text, value ?? string.Empty));
            }

            return new ResultBuffer(typedValues.ToArray());
        }

        public static string ReadText(TypedValue[] values, int index)
        {
            if (values == null || index < 0 || index >= values.Length)
            {
                return string.Empty;
            }

            return values[index].Value == null
                ? string.Empty
                : values[index].Value.ToString();
        }

        private static DBDictionary GetOrCreateAppDictionary(Entity entity, Transaction tr)
        {
            if (entity.ExtensionDictionary.IsNull)
            {
                entity.UpgradeOpen();
                entity.CreateExtensionDictionary();
            }

            var extensionDictionary = (DBDictionary)tr.GetObject(entity.ExtensionDictionary, OpenMode.ForWrite);
            if (extensionDictionary.Contains(DictionaryName))
            {
                return (DBDictionary)tr.GetObject(extensionDictionary.GetAt(DictionaryName), OpenMode.ForWrite);
            }

            var appDictionary = new DBDictionary();
            extensionDictionary.SetAt(DictionaryName, appDictionary);
            tr.AddNewlyCreatedDBObject(appDictionary, true);
            return appDictionary;
        }
    }
}
