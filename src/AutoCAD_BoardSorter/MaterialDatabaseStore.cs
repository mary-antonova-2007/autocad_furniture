using System;
using System.IO;
using System.Runtime.Serialization.Json;
using AutoCAD_BoardSorter.Models;

namespace AutoCAD_BoardSorter
{
    internal static class MaterialDatabaseStore
    {
        private const string FolderName = "AutoCAD_BoardSorter";
        private const string FileName = "materials.json";

        public static string FilePath
        {
            get
            {
                string root = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
                return Path.Combine(root, FolderName, FileName);
            }
        }

        public static MaterialDatabase Load()
        {
            try
            {
                string path = FilePath;
                if (!File.Exists(path))
                {
                    MaterialDatabase seeded = CreateDefault();
                    Save(seeded);
                    return seeded;
                }

                using (var stream = File.OpenRead(path))
                {
                    var serializer = new DataContractJsonSerializer(typeof(MaterialDatabase));
                    var db = serializer.ReadObject(stream) as MaterialDatabase;
                    return Normalize(db);
                }
            }
            catch
            {
                return CreateDefault();
            }
        }

        public static void Save(MaterialDatabase database)
        {
            string path = FilePath;
            Directory.CreateDirectory(Path.GetDirectoryName(path));
            using (var stream = File.Create(path))
            {
                var serializer = new DataContractJsonSerializer(typeof(MaterialDatabase));
                serializer.WriteObject(stream, Normalize(database));
            }
        }

        public static string DisplayName(BoardMaterialData material)
        {
            if (material == null)
            {
                return string.Empty;
            }

            string code = (material.Code ?? string.Empty).Trim();
            string name = (material.Name ?? string.Empty).Trim();
            if (string.IsNullOrWhiteSpace(code))
            {
                return name;
            }

            if (string.IsNullOrWhiteSpace(name))
            {
                return code;
            }

            return code + " " + name;
        }

        public static string DisplayName(CoatingMaterialData material)
        {
            if (material == null)
            {
                return string.Empty;
            }

            string code = (material.Code ?? string.Empty).Trim();
            string name = (material.Name ?? string.Empty).Trim();
            if (string.IsNullOrWhiteSpace(code))
            {
                return name;
            }

            if (string.IsNullOrWhiteSpace(name))
            {
                return code;
            }

            return code + " " + name;
        }

        private static MaterialDatabase Normalize(MaterialDatabase db)
        {
            if (db == null)
            {
                db = new MaterialDatabase();
            }

            if (db.BoardMaterials == null) db.BoardMaterials = new System.Collections.Generic.List<BoardMaterialData>();
            if (db.CoatingMaterials == null) db.CoatingMaterials = new System.Collections.Generic.List<CoatingMaterialData>();
            if (db.BoardCategories == null) db.BoardCategories = new System.Collections.Generic.List<MaterialCategoryData>();
            if (db.CoatingCategories == null) db.CoatingCategories = new System.Collections.Generic.List<MaterialCategoryData>();
            if (db.UiLayout == null) db.UiLayout = new MaterialDatabaseUiLayout();
            foreach (BoardMaterialData material in db.BoardMaterials)
            {
                if (material.Formats == null)
                {
                    material.Formats = new System.Collections.Generic.List<MaterialFormatData>();
                }
            }

            return db;
        }

        private static MaterialDatabase CreateDefault()
        {
            var db = new MaterialDatabase();
            db.BoardCategories.Add(new MaterialCategoryData { Code = "sheet", Name = "Листовые материалы" });
            db.CoatingCategories.Add(new MaterialCategoryData { Code = "edge", Name = "Кромка" });
            db.CoatingCategories.Add(new MaterialCategoryData { Code = "face", Name = "Пласть" });
            db.CoatingMaterials.Add(new CoatingMaterialData
            {
                Code = "PVC-2",
                Name = "ПВХ 2 мм",
                CalculationType = "Погонный",
                Thickness = 2.0,
                CategoryCode = "edge"
            });
            db.CoatingMaterials.Add(new CoatingMaterialData
            {
                Code = "PVC-05",
                Name = "ПВХ 0,5 мм",
                CalculationType = "Погонный",
                Thickness = 0.5,
                CategoryCode = "edge"
            });
            db.BoardMaterials.Add(new BoardMaterialData
            {
                Code = "LDSP-16",
                Name = "ЛДСП 16 мм",
                CalculationType = "Площадной",
                CategoryCode = "sheet",
                DefaultVisibleEdgeCoatingCode = "PVC-2",
                DefaultHiddenEdgeCoatingCode = "PVC-05",
                Formats =
                {
                    new MaterialFormatData
                    {
                        Code = "2800x2070x16",
                        FormatType = "Площадной",
                        Length = 2800,
                        Width = 2070,
                        Thickness = 16
                    }
                }
            });
            return db;
        }
    }
}
