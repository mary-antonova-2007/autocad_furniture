using System;
using System.Collections.Generic;
using Autodesk.AutoCAD.DatabaseServices;

namespace AutoCAD_BoardSorter
{
    internal static class FaceCoatingStorage
    {
        private const string DictionaryName = "AutoCAD_BoardSorter";
        private const string RecordName = "FaceCoatings";
        private const string OriginalColorRecordName = "FaceCoatingOriginalColors";
        private const int Version = 1;

        public static string Read(Entity entity, Transaction tr, string faceKey)
        {
            Dictionary<string, string> values = ReadFaceCoatings(entity, tr);
            string coating;
            return values.TryGetValue(faceKey, out coating) ? coating : string.Empty;
        }

        public static Dictionary<string, string> ReadFaceCoatings(Entity entity, Transaction tr)
        {
            return ReadStringMap(entity, tr, RecordName);
        }

        public static void Write(Entity entity, Transaction tr, string faceKey, string coating)
        {
            Dictionary<string, string> values = ReadFaceCoatings(entity, tr);

            if (string.IsNullOrWhiteSpace(coating))
            {
                values.Remove(faceKey);
            }
            else
            {
                values[faceKey] = coating;
            }

            WriteStringMap(entity, tr, RecordName, values);
        }

        public static Dictionary<string, string> ReadOriginalColors(Entity entity, Transaction tr)
        {
            return ReadStringMap(entity, tr, OriginalColorRecordName);
        }

        public static void WriteOriginalColors(Entity entity, Transaction tr, Dictionary<string, string> values)
        {
            WriteStringMap(entity, tr, OriginalColorRecordName, values);
        }

        public static void ClearOriginalColors(Entity entity, Transaction tr)
        {
            if (entity.ExtensionDictionary.IsNull)
            {
                return;
            }

            var extensionDictionary = (DBDictionary)tr.GetObject(entity.ExtensionDictionary, OpenMode.ForRead);
            if (!extensionDictionary.Contains(DictionaryName))
            {
                return;
            }

            var appDictionary = (DBDictionary)tr.GetObject(extensionDictionary.GetAt(DictionaryName), OpenMode.ForWrite);
            if (!appDictionary.Contains(OriginalColorRecordName))
            {
                return;
            }

            DBObject old = tr.GetObject(appDictionary.GetAt(OriginalColorRecordName), OpenMode.ForWrite);
            old.Erase();
        }

        private static Dictionary<string, string> ReadStringMap(Entity entity, Transaction tr, string recordName)
        {
            var values = new Dictionary<string, string>(StringComparer.Ordinal);

            if (entity.ExtensionDictionary.IsNull)
            {
                return values;
            }

            var extensionDictionary = (DBDictionary)tr.GetObject(entity.ExtensionDictionary, OpenMode.ForRead);
            if (!extensionDictionary.Contains(DictionaryName))
            {
                return values;
            }

            var appDictionary = (DBDictionary)tr.GetObject(extensionDictionary.GetAt(DictionaryName), OpenMode.ForRead);
            if (!appDictionary.Contains(recordName))
            {
                return values;
            }

            var record = (Xrecord)tr.GetObject(appDictionary.GetAt(recordName), OpenMode.ForRead);
            if (record.Data == null)
            {
                return values;
            }

            TypedValue[] data = record.Data.AsArray();
            for (int i = 1; i + 1 < data.Length; i += 2)
            {
                string key = Convert.ToString(data[i].Value);
                string value = Convert.ToString(data[i + 1].Value);
                if (!string.IsNullOrWhiteSpace(key))
                {
                    values[key] = value ?? string.Empty;
                }
            }

            return values;
        }

        private static void WriteStringMap(Entity entity, Transaction tr, string recordName, Dictionary<string, string> values)
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

            if (appDictionary.Contains(recordName))
            {
                DBObject old = tr.GetObject(appDictionary.GetAt(recordName), OpenMode.ForWrite);
                old.Erase();
            }

            var typedValues = new List<TypedValue>
            {
                new TypedValue((int)DxfCode.Int16, Version)
            };

            foreach (KeyValuePair<string, string> pair in values)
            {
                typedValues.Add(new TypedValue((int)DxfCode.Text, pair.Key));
                typedValues.Add(new TypedValue((int)DxfCode.Text, pair.Value ?? string.Empty));
            }

            var record = new Xrecord
            {
                Data = new ResultBuffer(typedValues.ToArray())
            };

            appDictionary.SetAt(recordName, record);
            tr.AddNewlyCreatedDBObject(record, true);
        }
    }
}
