using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Reflection;
using System.Text;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.BoundaryRepresentation;
using Autodesk.AutoCAD.Colors;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;
using AutoCAD_BoardSorter.Geometry;
using AutoCAD_BoardSorter.Models;
using AutoCAD_BoardSorter.Ui;
using WinRegistry = Microsoft.Win32.Registry;
using WinRegistryKey = Microsoft.Win32.RegistryKey;
using WinRegistryValueKind = Microsoft.Win32.RegistryValueKind;

[assembly: CommandClass(typeof(AutoCAD_BoardSorter.BoardSorterCommands))]

namespace AutoCAD_BoardSorter
{
    public sealed class BoardSorterCommands : IExtensionApplication
    {
        private const string AppName = "AutoCAD_BoardSorter";

        public void Initialize()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            if (doc != null)
            {
                doc.Editor.WriteMessage("\nAutoCAD_BoardSorter loaded. Commands: BDSORT, BDPALETTE, BDFACECOLORS, BDCLEARFACECOLORS, BDSPEC, BDSPECINFO, BDSPECCLEAR.");
            }
        }

        public void Terminate()
        {
        }

        [CommandMethod("BDSORT")]
        public void SortBoards()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            if (doc == null)
            {
                return;
            }

            Database db = doc.Database;
            Editor ed = doc.Editor;

            string assemblyNumber = PromptString(ed, "\nНомер сборки", true);
            if (assemblyNumber == null)
            {
                return;
            }

            var ids = PromptSolidSelection(ed);
            if (ids == null)
            {
                ids = GetAllSolidIds(db);
            }

            if (ids.Count == 0)
            {
                ed.WriteMessage("\nВ чертеже нет 3DSOLID.");
                return;
            }

            var analyzer = new BoardDimensionAnalyzer();
            var coatingAnalyzer = new BoardCoatingAnalyzer();
            var fingerprintBuilder = new SolidFingerprintBuilder();
            var boards = new List<BoardInfo>();
            string sortLogPath;

            using (Transaction tr = db.TransactionManager.StartTransaction())
            using (BoardSortLogger log = BoardSortLogger.Create(db, "bdsort"))
            {
                sortLogPath = log.Path;
                log.Info("BDSORT started. Input solids: " + ids.Count.ToString(CultureInfo.InvariantCulture));
                int number = 1;
                foreach (ObjectId id in ids)
                {
                    var solid = tr.GetObject(id, OpenMode.ForRead, false) as Solid3d;
                    if (solid == null || solid.IsErased)
                    {
                        continue;
                    }

                    log.Info("Solid " + solid.Handle + " layer=" + solid.Layer);

                    double length;
                    double width;
                    double thickness;
                    string method;

                    if (!analyzer.TryGetDimensions(solid, out length, out width, out thickness, out method))
                    {
                        ed.WriteMessage("\nПропуск {0}: не удалось посчитать габарит.", solid.Handle);
                        log.Info("  dimensions failed");
                        continue;
                    }

                    log.Info("  dimensions L=" + Round1(length).ToString("0.###", CultureInfo.InvariantCulture)
                        + " W=" + Round1(width).ToString("0.###", CultureInfo.InvariantCulture)
                        + " T=" + Round1(thickness).ToString("0.###", CultureInfo.InvariantCulture)
                        + " method=" + method);

                    SpecificationData existingSpecification;
                    SpecificationStorage.TryRead(solid, tr, out existingSpecification);

                    boards.Add(new BoardInfo
                    {
                        ObjectId = id,
                        Number = number++,
                        Handle = solid.Handle.ToString(),
                        Layer = solid.Layer,
                        LengthMm = Round1(length),
                        WidthMm = Round1(width),
                        ThicknessMm = Round1(thickness),
                        PartName = FirstNonBlank(existingSpecification == null ? null : existingSpecification.PartName, "Деталь"),
                        Material = FirstNonBlank(existingSpecification == null ? null : existingSpecification.Material, (string.IsNullOrWhiteSpace(solid.Layer) ? "Материал" : solid.Layer.Trim()) + " " + FormatMaterialThickness(Round1(thickness)) + " мм"),
                        Method = method,
                        RotateLengthWidth = existingSpecification != null && existingSpecification.RotateLengthWidth,
                        Coatings = coatingAnalyzer.Analyze(solid, tr, existingSpecification, log),
                        Sketch = coatingAnalyzer.BuildSketch(solid, tr, existingSpecification, log)
                    });

                    boards[boards.Count - 1].Fingerprint = fingerprintBuilder.Build(
                        solid,
                        boards[boards.Count - 1].LengthMm,
                        boards[boards.Count - 1].WidthMm,
                        boards[boards.Count - 1].ThicknessMm);
                }

                tr.Commit();
                log.Info("BDSORT finished boards=" + boards.Count.ToString(CultureInfo.InvariantCulture));
            }

            if (boards.Count == 0)
            {
                ed.WriteMessage("\nПодходящих тел не найдено.");
                return;
            }

            var groups = AggregateBoards(boards)
                .OrderBy(x => x.Layer, StringComparer.CurrentCultureIgnoreCase)
                .ThenBy(x => x.ThicknessMm)
                .ThenByDescending(x => x.LengthMm)
                .ThenByDescending(x => x.WidthMm)
                .ToList();

            for (int i = 0; i < groups.Count; i++)
            {
                groups[i].Number = i + 1;
            }

            int specificationCount = WriteSpecifications(db, groups, assemblyNumber);

            PrintReport(ed, groups);
            string csvPath = WriteCsv(db, groups, assemblyNumber);
            string xlsxPath = WriteXlsx(db, groups, assemblyNumber);
            ed.WriteMessage("\nСпецификация записана в тел: {0}.", specificationCount);
            ed.WriteMessage("\nCSV: {0}", csvPath);
            ed.WriteMessage("\nExcel: {0}", xlsxPath);
            ed.WriteMessage("\nЛог BDSORT: {0}", sortLogPath);
        }

        [CommandMethod("BDPALETTE")]
        public void ShowSpecificationPalette()
        {
            SpecificationPalette.Show();
        }

        [CommandMethod("BDFACECOLORS")]
        public void ShowFaceCoatingColors()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            if (doc == null)
            {
                return;
            }

            Database db = doc.Database;
            Editor ed = doc.Editor;

            List<ObjectId> ids = PromptSolidSelection(ed);
            if (ids == null)
            {
                ids = GetAllSolidIds(db);
            }

            using (Transaction tr = db.TransactionManager.StartTransaction())
            using (BoardSortLogger log = BoardSortLogger.Create(db, "facecolors"))
            {
                int colored = 0;
                int solids = 0;
                int coatingRecords = 0;
                int resolved = 0;
                int unresolved = 0;
                int colorErrors = 0;

                log.Info("BDFACECOLORS started. Input solids: " + ids.Count.ToString(CultureInfo.InvariantCulture));

                foreach (ObjectId id in ids)
                {
                    var solid = tr.GetObject(id, OpenMode.ForWrite, false) as Solid3d;
                    if (solid == null || solid.IsErased)
                    {
                        continue;
                    }

                    Dictionary<string, string> coatings = FaceCoatingStorage.ReadFaceCoatings(solid, tr);
                    log.Info("Solid " + solid.Handle + " layer=" + solid.Layer + " coatings=" + coatings.Count.ToString(CultureInfo.InvariantCulture));
                    if (coatings.Count == 0)
                    {
                        continue;
                    }

                    Dictionary<string, string> originals = FaceCoatingStorage.ReadOriginalColors(solid, tr);
                    bool originalsChanged = false;
                    var processedFaceOrdinals = new HashSet<long>();

                    foreach (KeyValuePair<string, string> pair in coatings)
                    {
                        if (string.IsNullOrWhiteSpace(pair.Value))
                        {
                            log.Info("  key=" + pair.Key + " skipped empty coating");
                            continue;
                        }

                        coatingRecords++;
                        log.Info("  key=" + pair.Key + " coating=\"" + pair.Value + "\"");

                        SubentityId subentityId;
                        if (!TryResolveFaceSubentityId(solid, pair.Key, out subentityId, log))
                        {
                            unresolved++;
                            log.Info("    unresolved");
                            continue;
                        }

                        resolved++;
                        log.Info("    resolved type=" + subentityId.Type + " index=" + subentityId.IndexPtr.ToInt64().ToString(CultureInfo.InvariantCulture));
                        if (!processedFaceOrdinals.Add(subentityId.IndexPtr.ToInt64()))
                        {
                            log.Info("    skipped duplicate resolved face");
                            continue;
                        }

                        try
                        {
                            if (!originals.ContainsKey(pair.Key))
                            {
                                try
                                {
                                    Color original = solid.GetSubentityColor(subentityId);
                                    originals[pair.Key] = SerializeColor(original);
                                    log.Info("    original color=" + originals[pair.Key]);
                                    originalsChanged = true;
                                }
                                catch (Autodesk.AutoCAD.Runtime.Exception ex)
                                {
                                    log.Error("    original color read failed " + ex.ErrorStatus, ex);
                                    originals[pair.Key] = SerializeColor(Color.FromColorIndex(ColorMethod.ByLayer, 256));
                                    log.Info("    original color fallback=ByLayer");
                                    originalsChanged = true;
                                }
                                catch (System.Exception ex)
                                {
                                    log.Error("    original color read failed", ex);
                                    originals[pair.Key] = SerializeColor(Color.FromColorIndex(ColorMethod.ByLayer, 256));
                                    log.Info("    original color fallback=ByLayer");
                                    originalsChanged = true;
                                }
                            }

                            Color targetColor = ColorForCoating(pair.Value);
                            log.Info("    target color=" + SerializeColor(targetColor));
                            solid.SetSubentityColor(subentityId, targetColor);
                            colored++;
                            log.Info("    colored ok");
                        }
                        catch (Autodesk.AutoCAD.Runtime.Exception ex)
                        {
                            colorErrors++;
                            ed.WriteMessage("\nГрань {0} пропущена: {1}.", pair.Key, ex.ErrorStatus);
                            log.Error("    AutoCAD exception " + ex.ErrorStatus, ex);
                        }
                        catch (System.Exception ex)
                        {
                            colorErrors++;
                            ed.WriteMessage("\nГрань {0} пропущена: {1}.", pair.Key, ex.Message);
                            log.Error("    .NET exception", ex);
                        }
                    }

                    if (originalsChanged)
                    {
                        FaceCoatingStorage.WriteOriginalColors(solid, tr, originals);
                    }

                    solids++;
                }

                tr.Commit();
                log.Info("Summary records=" + coatingRecords.ToString(CultureInfo.InvariantCulture)
                    + " resolved=" + resolved.ToString(CultureInfo.InvariantCulture)
                    + " unresolved=" + unresolved.ToString(CultureInfo.InvariantCulture)
                    + " colored=" + colored.ToString(CultureInfo.InvariantCulture)
                    + " colorErrors=" + colorErrors.ToString(CultureInfo.InvariantCulture)
                    + " solids=" + solids.ToString(CultureInfo.InvariantCulture));
                ed.WriteMessage("\nРаскрашено граней: {0}, тел: {1}. Лог: {2}", colored, solids, log.Path);
            }
        }

        [CommandMethod("BDCLEARFACECOLORS")]
        public void ClearFaceCoatingColors()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            if (doc == null)
            {
                return;
            }

            Database db = doc.Database;
            Editor ed = doc.Editor;

            List<ObjectId> ids = PromptSolidSelection(ed);
            if (ids == null)
            {
                ids = GetAllSolidIds(db);
            }

            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                int restored = 0;

                foreach (ObjectId id in ids)
                {
                    var solid = tr.GetObject(id, OpenMode.ForWrite, false) as Solid3d;
                    if (solid == null || solid.IsErased)
                    {
                        continue;
                    }

                    Dictionary<string, string> originals = FaceCoatingStorage.ReadOriginalColors(solid, tr);
                    foreach (KeyValuePair<string, string> pair in originals)
                    {
                        SubentityId subentityId;
                        Color color;
                        if (!TryResolveFaceSubentityId(solid, pair.Key, out subentityId) || !TryDeserializeColor(pair.Value, out color))
                        {
                            continue;
                        }

                        try
                        {
                            solid.SetSubentityColor(subentityId, color);
                            restored++;
                        }
                        catch (Autodesk.AutoCAD.Runtime.Exception ex)
                        {
                            ed.WriteMessage("\nГрань {0} не восстановлена: {1}.", pair.Key, ex.ErrorStatus);
                        }
                    }

                    FaceCoatingStorage.ClearOriginalColors(solid, tr);
                }

                tr.Commit();
                ed.WriteMessage("\nВосстановлено цветов граней: {0}.", restored);
            }
        }

        [CommandMethod("BDSPEC")]
        public void WriteSpecification()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            if (doc == null)
            {
                return;
            }

            Editor ed = doc.Editor;
            Database db = doc.Database;

            List<ObjectId> ids = PromptSolidSelection(ed);
            if (ids == null || ids.Count == 0)
            {
                ed.WriteMessage("\nНичего не выбрано.");
                return;
            }

            SpecificationData data = PromptSpecificationData(ed);
            if (data == null)
            {
                return;
            }

            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                int written = 0;
                foreach (ObjectId id in ids)
                {
                    var solid = tr.GetObject(id, OpenMode.ForWrite, false) as Solid3d;
                    if (solid == null || solid.IsErased)
                    {
                        continue;
                    }

                    SpecificationStorage.Write(solid, data, tr);
                    written++;
                }

                tr.Commit();
                ed.WriteMessage("\nСпецификация записана в тел: {0}.", written);
            }
        }

        [CommandMethod("BDSPECINFO")]
        public void ShowSpecification()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            if (doc == null)
            {
                return;
            }

            Editor ed = doc.Editor;
            Database db = doc.Database;

            PromptEntityOptions options = new PromptEntityOptions("\nВыбери 3D-тело: ");
            options.SetRejectMessage("\nНужен 3DSOLID.");
            options.AddAllowedClass(typeof(Solid3d), false);
            PromptEntityResult result = ed.GetEntity(options);
            if (result.Status != PromptStatus.OK)
            {
                return;
            }

            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                var solid = (Solid3d)tr.GetObject(result.ObjectId, OpenMode.ForRead);
                SpecificationData data;
                if (!SpecificationStorage.TryRead(solid, tr, out data))
                {
                    ed.WriteMessage("\nУ тела нет параметров Спецификация.");
                    return;
                }

                PrintSpecification(ed, data);
                tr.Commit();
            }
        }

        [CommandMethod("BDSPECCLEAR")]
        public void ClearSpecification()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            if (doc == null)
            {
                return;
            }

            Editor ed = doc.Editor;
            Database db = doc.Database;

            List<ObjectId> ids = PromptSolidSelection(ed);
            if (ids == null || ids.Count == 0)
            {
                ed.WriteMessage("\nНичего не выбрано.");
                return;
            }

            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                int cleared = 0;
                foreach (ObjectId id in ids)
                {
                    var solid = tr.GetObject(id, OpenMode.ForRead, false) as Solid3d;
                    if (solid == null || solid.IsErased)
                    {
                        continue;
                    }

                    if (SpecificationStorage.Clear(solid, tr))
                    {
                        cleared++;
                    }
                }

                tr.Commit();
                ed.WriteMessage("\nСпецификация очищена у тел: {0}.", cleared);
            }
        }

        [CommandMethod("BDREGISTER")]
        public void RegisterAutoLoad()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;

            try
            {
                string productKey = GetRegistryProductRootKey();
                if (string.IsNullOrWhiteSpace(productKey))
                {
                    ed.WriteMessage("\nНе удалось определить ветку реестра AutoCAD. Используй install-bundle.bat для автозагрузки.");
                    return;
                }

                using (WinRegistryKey product = WinRegistry.CurrentUser.OpenSubKey(productKey))
                using (WinRegistryKey applications = product.OpenSubKey("Applications", true))
                using (WinRegistryKey app = applications.CreateSubKey(AppName))
                {
                    string assemblyPath = Assembly.GetExecutingAssembly().Location;
                    app.SetValue("DESCRIPTION", "Board dimensions sorter by layer and thickness", WinRegistryValueKind.String);
                    app.SetValue("LOADCTRLS", 14, WinRegistryValueKind.DWord);
                    app.SetValue("LOADER", assemblyPath, WinRegistryValueKind.String);
                    app.SetValue("MANAGED", 1, WinRegistryValueKind.DWord);
                }

                ed.WriteMessage("\n{0} зарегистрирован для автозагрузки. Перезапусти AutoCAD.", AppName);
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage("\nНе удалось зарегистрировать автозагрузку: {0}", ex.Message);
            }
        }

        [CommandMethod("BDUNREGISTER")]
        public void UnregisterAutoLoad()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;

            try
            {
                string productKey = GetRegistryProductRootKey();
                if (string.IsNullOrWhiteSpace(productKey))
                {
                    ed.WriteMessage("\nНе удалось определить ветку реестра AutoCAD.");
                    return;
                }

                using (WinRegistryKey product = WinRegistry.CurrentUser.OpenSubKey(productKey))
                using (WinRegistryKey applications = product.OpenSubKey("Applications", true))
                {
                    applications.DeleteSubKeyTree(AppName, false);
                }

                ed.WriteMessage("\n{0} удален из автозагрузки. Перезапусти AutoCAD.", AppName);
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage("\nНе удалось удалить автозагрузку: {0}", ex.Message);
            }
        }

        private static string GetRegistryProductRootKey()
        {
            object value = TryGetStaticProperty(typeof(Application), "UserRegistryProductRootKey")
                ?? TryGetStaticProperty(typeof(Application), "MachineRegistryProductRootKey")
                ?? TryGetInstanceProperty(HostApplicationServices.Current, "RegistryProductRootKey")
                ?? TryGetInstanceProperty(HostApplicationServices.Current, "UserRegistryProductRootKey");

            return value as string;
        }

        private static object TryGetStaticProperty(Type type, string propertyName)
        {
            try
            {
                PropertyInfo property = type.GetProperty(propertyName, BindingFlags.Public | BindingFlags.Static);
                return property == null ? null : property.GetValue(null, null);
            }
            catch
            {
                return null;
            }
        }

        private static object TryGetInstanceProperty(object instance, string propertyName)
        {
            if (instance == null)
            {
                return null;
            }

            try
            {
                PropertyInfo property = instance.GetType().GetProperty(propertyName, BindingFlags.Public | BindingFlags.Instance);
                return property == null ? null : property.GetValue(instance, null);
            }
            catch
            {
                return null;
            }
        }

        private static List<ObjectId> PromptSolidSelection(Editor ed)
        {
            var options = new PromptSelectionOptions
            {
                MessageForAdding = "\nВыбери 3D-тела или нажми Enter для всех: "
            };

            var filter = new SelectionFilter(new[]
            {
                new TypedValue((int)DxfCode.Start, "3DSOLID")
            });

            PromptSelectionResult result = ed.GetSelection(options, filter);
            if (result.Status == PromptStatus.OK)
            {
                return result.Value.GetObjectIds().ToList();
            }

            if (result.Status == PromptStatus.None || result.Status == PromptStatus.Cancel)
            {
                return null;
            }

            return null;
        }

        private static List<ObjectId> GetAllSolidIds(Database db)
        {
            var ids = new List<ObjectId>();

            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                var blockTable = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
                var modelSpace = (BlockTableRecord)tr.GetObject(blockTable[BlockTableRecord.ModelSpace], OpenMode.ForRead);

                foreach (ObjectId id in modelSpace)
                {
                    if (id.ObjectClass != RXObject.GetClass(typeof(Solid3d)))
                    {
                        continue;
                    }

                    ids.Add(id);
                }

                tr.Commit();
            }

            return ids;
        }

        private static SpecificationData PromptSpecificationData(Editor ed)
        {
            string assemblyNumber = PromptString(ed, "\nНомер сборки", true);
            if (assemblyNumber == null) return null;

            string partNumber = PromptString(ed, "\nНомер детали", true);
            if (partNumber == null) return null;

            string partName = PromptString(ed, "\nНаименование детали", true);
            if (partName == null) return null;

            string partType = PromptPartType(ed);
            if (partType == null) return null;

            double length = PromptDouble(ed, "\nДлина", 0.0);
            if (double.IsNaN(length)) return null;

            double width = PromptDouble(ed, "\nШирина", 0.0);
            if (double.IsNaN(width)) return null;

            bool? rotate = PromptBool(ed, "\nПоворот, поменять длину и ширину местами");
            if (!rotate.HasValue) return null;

            string material = PromptString(ed, "\nМатериал", true);
            if (material == null) return null;

            string note = PromptString(ed, "\nПримечание", true);
            if (note == null) return null;

            if (rotate.Value)
            {
                double t = length;
                length = width;
                width = t;
            }

            return new SpecificationData
            {
                AssemblyNumber = assemblyNumber,
                PartNumber = partNumber,
                PartName = partName,
                PartType = partType,
                LengthMm = length,
                WidthMm = width,
                RotateLengthWidth = rotate.Value,
                Material = material,
                Note = note
            };
        }

        private static string PromptPartType(Editor ed)
        {
            var options = new PromptKeywordOptions("\nТип детали")
            {
                AllowNone = false
            };
            options.Keywords.Add("Погонный");
            options.Keywords.Add("Площадной");
            options.Keywords.Add("Объёмный");
            options.Keywords.Default = "Площадной";

            PromptResult result = ed.GetKeywords(options);
            return result.Status == PromptStatus.OK ? result.StringResult : null;
        }

        private static string PromptString(Editor ed, string message, bool allowSpaces)
        {
            var options = new PromptStringOptions(message + ": ")
            {
                AllowSpaces = allowSpaces
            };

            PromptResult result = ed.GetString(options);
            if (result.Status == PromptStatus.Cancel)
            {
                return null;
            }

            return result.Status == PromptStatus.OK ? result.StringResult : string.Empty;
        }

        private static double PromptDouble(Editor ed, string message, double defaultValue)
        {
            var options = new PromptDoubleOptions(message + " <" + Format(defaultValue) + ">: ")
            {
                AllowNone = true,
                AllowNegative = false,
                AllowZero = true,
                DefaultValue = defaultValue
            };

            PromptDoubleResult result = ed.GetDouble(options);
            if (result.Status == PromptStatus.Cancel)
            {
                return double.NaN;
            }

            return result.Status == PromptStatus.OK ? result.Value : defaultValue;
        }

        private static bool? PromptBool(Editor ed, string message)
        {
            var options = new PromptKeywordOptions(message + " [Да/Нет] <Нет>: ")
            {
                AllowNone = true
            };
            options.Keywords.Add("Да");
            options.Keywords.Add("Нет");
            options.Keywords.Default = "Нет";

            PromptResult result = ed.GetKeywords(options);
            if (result.Status == PromptStatus.Cancel)
            {
                return null;
            }

            return result.Status == PromptStatus.OK && result.StringResult == "Да";
        }

        private static void PrintSpecification(Editor ed, SpecificationData data)
        {
            ed.WriteMessage("\nСпецификация:");
            ed.WriteMessage("\n  Номер сборки: {0}", data.AssemblyNumber);
            ed.WriteMessage("\n  Номер детали: {0}", data.PartNumber);
            ed.WriteMessage("\n  Наименование детали: {0}", data.PartName);
            ed.WriteMessage("\n  Тип детали: {0}", data.PartType);
            ed.WriteMessage("\n  Длина: {0:0.0}", data.LengthMm);
            ed.WriteMessage("\n  Ширина: {0:0.0}", data.WidthMm);
            ed.WriteMessage("\n  Поворот: {0}", data.RotateLengthWidth ? "Да" : "Нет");
            ed.WriteMessage("\n  Материал: {0}", data.Material);
            ed.WriteMessage("\n  Примечание: {0}", data.Note);
        }

        private static List<BoardGroup> AggregateBoards(IList<BoardInfo> boards)
        {
            return boards
                .GroupBy(x => new
                {
                    x.Layer,
                    x.LengthMm,
                    x.WidthMm,
                    x.ThicknessMm,
                    x.Fingerprint,
                    x.RotateLengthWidth,
                    x.PartName,
                    x.Material,
                    P1 = Slot(x.Coatings, "P1"),
                    P2 = Slot(x.Coatings, "P2"),
                    L1 = Slot(x.Coatings, "L1"),
                    L2 = Slot(x.Coatings, "L2"),
                    W1 = Slot(x.Coatings, "W1"),
                    W2 = Slot(x.Coatings, "W2")
                })
                .Select(g =>
                {
                    var first = g.First();
                    var boardGroup = new BoardGroup
                    {
                        Layer = first.Layer,
                        LengthMm = first.LengthMm,
                        WidthMm = first.WidthMm,
                        ThicknessMm = first.ThicknessMm,
                        PartName = first.PartName,
                        Material = first.Material,
                        Method = first.Method,
                        RotateLengthWidth = first.RotateLengthWidth,
                        Coatings = first.Coatings ?? new BoardCoatingSlots(),
                        Sketch = first.Sketch
                    };
                    boardGroup.Items.AddRange(g.OrderBy(x => x.Number));
                    return boardGroup;
                })
                .ToList();
        }

        private static string Slot(BoardCoatingSlots coatings, string name)
        {
            if (coatings == null)
            {
                return string.Empty;
            }

            switch (name)
            {
                case "P1": return NormalizePlateSlot(coatings, true);
                case "P2": return NormalizePlateSlot(coatings, false);
                case "L1": return coatings.L1 ?? string.Empty;
                case "L2": return coatings.L2 ?? string.Empty;
                case "W1": return coatings.W1 ?? string.Empty;
                case "W2": return coatings.W2 ?? string.Empty;
                default: return string.Empty;
            }
        }

        private static string NormalizePlateSlot(BoardCoatingSlots coatings, bool first)
        {
            string primary = (coatings.P1 ?? string.Empty).Trim();
            string secondary = (coatings.P2 ?? string.Empty).Trim();
            if (!string.IsNullOrWhiteSpace(primary) && !string.IsNullOrWhiteSpace(secondary))
            {
                return first ? primary : secondary;
            }

            if (!string.IsNullOrWhiteSpace(primary))
            {
                return primary;
            }

            if (!string.IsNullOrWhiteSpace(secondary))
            {
                return secondary;
            }

            return string.Empty;
        }

        private static int WriteSpecifications(Database db, IList<BoardGroup> groups, string assemblyNumber)
        {
            int written = 0;

            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                foreach (BoardGroup group in groups)
                {
                    foreach (BoardInfo board in group.Items)
                    {
                        var solid = tr.GetObject(board.ObjectId, OpenMode.ForWrite, false) as Solid3d;
                        if (solid == null || solid.IsErased)
                        {
                            continue;
                        }

                        SpecificationData existing;
                        SpecificationStorage.TryRead(solid, tr, out existing);

                        SpecificationStorage.Write(solid, BuildSpecification(group, existing, assemblyNumber), tr);
                        written++;
                    }
                }

                tr.Commit();
            }

            return written;
        }

        private static SpecificationData BuildSpecification(BoardGroup group, SpecificationData existing, string assemblyNumber)
        {
            bool rotate = existing != null ? existing.RotateLengthWidth : group.RotateLengthWidth;
            double length = rotate ? group.WidthMm : group.LengthMm;
            double width = rotate ? group.LengthMm : group.WidthMm;
            string generatedNote = BuildGeneratedNote(group);
            string existingNote = existing == null ? string.Empty : (existing.Note ?? string.Empty);
            string finalNote = string.IsNullOrWhiteSpace(existingNote)
                ? generatedNote
                : (string.IsNullOrWhiteSpace(generatedNote) ? existingNote : existingNote + "; " + generatedNote);

            return new SpecificationData
            {
                AssemblyNumber = assemblyNumber ?? string.Empty,
                PartNumber = group.Number.ToString(CultureInfo.InvariantCulture),
                PartName = FirstNonBlank(existing == null ? null : existing.PartName, "Деталь"),
                PartType = FirstNonBlank(existing == null ? null : existing.PartType, "Площадной"),
                LengthMm = length,
                WidthMm = width,
                RotateLengthWidth = rotate,
                Material = FirstNonBlank(existing == null ? null : existing.Material, DefaultMaterial(group)),
                Note = finalNote
            };
        }

        private static string DefaultMaterial(BoardGroup group)
        {
            string layer = string.IsNullOrWhiteSpace(group.Layer) ? "Материал" : group.Layer.Trim();
            return layer + " " + FormatMaterialThickness(group.ThicknessMm) + " мм";
        }

        private static string DisplayPartName(BoardGroup group)
        {
            return group == null ? "Деталь" : FirstNonBlank(group.PartName, "Деталь");
        }

        private static string DisplayMaterial(BoardGroup group)
        {
            return group == null ? string.Empty : FirstNonBlank(group.Material, DefaultMaterial(group));
        }

        private static double DisplayLength(BoardGroup group)
        {
            return group != null && group.RotateLengthWidth ? group.WidthMm : group.LengthMm;
        }

        private static double DisplayWidth(BoardGroup group)
        {
            return group != null && group.RotateLengthWidth ? group.LengthMm : group.WidthMm;
        }

        private static string FormatMaterialThickness(double value)
        {
            double roundedInteger = Math.Round(value);
            if (Math.Abs(value - roundedInteger) < 0.05)
            {
                return roundedInteger.ToString("0", CultureInfo.InvariantCulture);
            }

            return value.ToString("0.#", CultureInfo.InvariantCulture);
        }

        private static string FirstNonBlank(string value, string fallback)
        {
            return string.IsNullOrWhiteSpace(value) ? fallback : value;
        }

        private static bool TryParseFaceKey(string key, out SubentityId subentityId)
        {
            subentityId = SubentityId.Null;

            if (string.IsNullOrWhiteSpace(key))
            {
                return false;
            }

            string[] parts = key.Split(':');
            if (parts.Length < 3)
            {
                return false;
            }

            long indexPtr;
            if (!long.TryParse(parts[2], NumberStyles.Integer, CultureInfo.InvariantCulture, out indexPtr))
            {
                return false;
            }

            subentityId = new SubentityId(SubentityType.Face, new IntPtr(indexPtr));
            return true;
        }

        private static bool TryResolveFaceSubentityId(Solid3d solid, string key, out SubentityId subentityId)
        {
            return TryResolveFaceSubentityId(solid, key, out subentityId, null);
        }

        private static bool TryResolveFaceSubentityId(Solid3d solid, string key, out SubentityId subentityId, BoardSortLogger log)
        {
            if (FaceKeyBuilder.IsFingerprintKey(key))
            {
                return TryResolveFaceByFingerprint(solid, key, out subentityId, log);
            }

            if (TryParseFaceKey(key, out subentityId))
            {
                if (log != null) log.Info("    legacy key resolved directly index=" + subentityId.IndexPtr.ToInt64().ToString(CultureInfo.InvariantCulture));
                return true;
            }

            string[] parts = (key ?? string.Empty).Split(':');
            if (parts.Length < 3)
            {
                return false;
            }

            long originalIndexPtr;
            if (!long.TryParse(parts[2], NumberStyles.Integer, CultureInfo.InvariantCulture, out originalIndexPtr))
            {
                return false;
            }

            try
            {
                using (var brep = new Brep(solid))
                {
                    int index = 0;
                    foreach (Autodesk.AutoCAD.BoundaryRepresentation.Face face in brep.Faces)
                    {
                        index++;
                        if (index == originalIndexPtr)
                        {
                            subentityId = new SubentityId(SubentityType.Face, new IntPtr(index));
                            if (log != null) log.Info("    legacy key matched brep ordinal " + index.ToString(CultureInfo.InvariantCulture));
                            return true;
                        }
                    }
                }
            }
            catch
            {
                return false;
            }

            return false;
        }

        private static bool TryResolveFaceByFingerprint(Solid3d solid, string key, out SubentityId subentityId, BoardSortLogger log)
        {
            subentityId = SubentityId.Null;

            try
            {
                using (var brep = new Brep(solid))
                {
                    int index = 0;
                    foreach (Autodesk.AutoCAD.BoundaryRepresentation.Face face in brep.Faces)
                    {
                        index++;
                        string candidateKey;
                        if (!FaceKeyBuilder.TryBuild(face, out candidateKey))
                        {
                            if (log != null) log.Info("    brep face " + index.ToString(CultureInfo.InvariantCulture) + " key build failed");
                            continue;
                        }

                        if (log != null) log.Info("    brep face " + index.ToString(CultureInfo.InvariantCulture) + " candidate=" + candidateKey);

                        if (candidateKey == key)
                        {
                            subentityId = new SubentityId(SubentityType.Face, new IntPtr(index));
                            if (log != null) log.Info("    fingerprint matched face ordinal " + index.ToString(CultureInfo.InvariantCulture));
                            return true;
                        }
                    }
                }
            }
            catch (System.Exception ex)
            {
                if (log != null) log.Error("    fingerprint resolve failed", ex);
                return false;
            }

            return false;
        }

        private static bool CanReadSubentityColor(Solid3d solid, SubentityId subentityId)
        {
            try
            {
                solid.GetSubentityColor(subentityId);
                return true;
            }
            catch
            {
                return false;
            }
        }

        private static Color ColorForCoating(string coating)
        {
            short[] aciColors = { 1, 3, 5, 2, 6, 4, 30, 70, 140, 200, 220, 240 };
            int hash = 17;
            string value = coating ?? string.Empty;
            for (int i = 0; i < value.Length; i++)
            {
                hash = unchecked(hash * 31 + value[i]);
            }

            short aci = aciColors[Math.Abs(hash) % aciColors.Length];
            return Color.FromColorIndex(ColorMethod.ByAci, aci);
        }

        private static string SerializeColor(Color color)
        {
            if (color == null || color.IsByLayer)
            {
                return "ByLayer";
            }

            if (color.IsByBlock)
            {
                return "ByBlock";
            }

            if (color.IsByAci)
            {
                return "Aci:" + color.ColorIndex.ToString(CultureInfo.InvariantCulture);
            }

            return "Rgb:"
                + color.Red.ToString(CultureInfo.InvariantCulture)
                + ","
                + color.Green.ToString(CultureInfo.InvariantCulture)
                + ","
                + color.Blue.ToString(CultureInfo.InvariantCulture);
        }

        private static bool TryDeserializeColor(string value, out Color color)
        {
            color = Color.FromColorIndex(ColorMethod.ByLayer, 256);

            if (string.IsNullOrWhiteSpace(value) || value == "ByLayer")
            {
                return true;
            }

            if (value == "ByBlock")
            {
                color = Color.FromColorIndex(ColorMethod.ByBlock, 0);
                return true;
            }

            if (value.StartsWith("Aci:", StringComparison.Ordinal))
            {
                short index;
                if (!short.TryParse(value.Substring(4), NumberStyles.Integer, CultureInfo.InvariantCulture, out index))
                {
                    return false;
                }

                color = Color.FromColorIndex(ColorMethod.ByAci, index);
                return true;
            }

            if (value.StartsWith("Rgb:", StringComparison.Ordinal))
            {
                string[] parts = value.Substring(4).Split(',');
                byte r;
                byte g;
                byte b;
                if (parts.Length != 3
                    || !byte.TryParse(parts[0], NumberStyles.Integer, CultureInfo.InvariantCulture, out r)
                    || !byte.TryParse(parts[1], NumberStyles.Integer, CultureInfo.InvariantCulture, out g)
                    || !byte.TryParse(parts[2], NumberStyles.Integer, CultureInfo.InvariantCulture, out b))
                {
                    return false;
                }

                color = Color.FromRgb(r, g, b);
                return true;
            }

            return false;
        }

        private static void PrintReport(Editor ed, IList<BoardGroup> groups)
        {
            ed.WriteMessage("\n{0}", new string('=', 132));
            ed.WriteMessage("\n{0,-4} {1,-22} {2,5} {3,10} {4,10} {5,12} {6,-30} {7}",
                "N", "Layer", "Qty", "Length", "Width", "Thickness", "Handles", "Method");
            ed.WriteMessage("\n{0}", new string('-', 132));

            foreach (BoardGroup group in groups)
            {
                ed.WriteMessage(
                    "\n{0,-4} {1,-22} {2,5} {3,10:0.0} {4,10:0.0} {5,12:0.0} {6,-30} {7}",
                    group.Number,
                    TrimForColumn(group.Layer, 22),
                    group.Quantity,
                    group.LengthMm,
                    group.WidthMm,
                    group.ThicknessMm,
                    TrimForColumn(group.Handles, 30),
                    group.Method);
            }

            ed.WriteMessage("\n{0}", new string('-', 132));
            ed.WriteMessage("\nГруппы по слой + толщина:");

            foreach (var group in groups.GroupBy(x => new { x.Layer, ThicknessKey = Format(x.ThicknessMm) }).OrderBy(g => g.Key.Layer).ThenBy(g => double.Parse(g.Key.ThicknessKey, CultureInfo.InvariantCulture)))
            {
                ed.WriteMessage("\n  {0} / {1} мм: {2} типов, {3} шт.", group.Key.Layer, group.Key.ThicknessKey, group.Count(), group.Sum(x => x.Quantity));
            }

            PrintMaterialSummary(ed, groups);
            PrintEdgeSummary(ed, groups);

            ed.WriteMessage("\n{0}", new string('=', 132));
        }

        private static string WriteCsv(Database db, IList<BoardGroup> groups, string assemblyNumber)
        {
            string drawingPath = db.Filename;
            string directory = string.IsNullOrWhiteSpace(drawingPath)
                ? Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory)
                : Path.GetDirectoryName(drawingPath);

            string name = string.IsNullOrWhiteSpace(drawingPath)
                ? "BoardSort"
                : Path.GetFileNameWithoutExtension(drawingPath);

            string path = Path.Combine(directory, name + "_boards.csv");
            Dictionary<string, int> coatingNumbers = BuildCoatingNumbers(groups);

            var sb = new StringBuilder();
            sb.AppendLine("Сборка;Деталь;Наименование;Материал;Длина;Ширина;Толщина;Кол-во;P1;P2;L1;L2;W1;W2;Примечание");

            foreach (BoardGroup group in groups)
            {
                sb.Append(EscapeCsv(assemblyNumber)).Append(';')
                    .Append(group.Number.ToString(CultureInfo.InvariantCulture)).Append(';')
                    .Append(EscapeCsv(DisplayPartName(group))).Append(';')
                    .Append(EscapeCsv(DisplayMaterial(group))).Append(';')
                    .Append(Format(DisplayLength(group))).Append(';')
                    .Append(Format(DisplayWidth(group))).Append(';')
                    .Append(Format(group.ThicknessMm)).Append(';')
                    .Append(group.Quantity.ToString(CultureInfo.InvariantCulture)).Append(';')
                    .Append(EscapeCsv(CoatingNumberText(Slot(group.Coatings, "P1"), coatingNumbers))).Append(';')
                    .Append(EscapeCsv(CoatingNumberText(Slot(group.Coatings, "P2"), coatingNumbers))).Append(';')
                    .Append(EscapeCsv(EdgeCoatingNumberText(Slot(group.Coatings, "L1"), coatingNumbers))).Append(';')
                    .Append(EscapeCsv(EdgeCoatingNumberText(Slot(group.Coatings, "L2"), coatingNumbers))).Append(';')
                    .Append(EscapeCsv(EdgeCoatingNumberText(Slot(group.Coatings, "W1"), coatingNumbers))).Append(';')
                    .Append(EscapeCsv(EdgeCoatingNumberText(Slot(group.Coatings, "W2"), coatingNumbers))).Append(';')
                    .Append(EscapeCsv(BuildGeneratedNote(group))).AppendLine();
            }

            sb.AppendLine();
            sb.AppendLine("Материалы;Тип;Количество");
            foreach (string line in BuildMaterialSummaryLines(groups))
            {
                sb.AppendLine(EscapeCsv(line));
            }

            sb.AppendLine();
            sb.AppendLine("Кромка;Количество");
            foreach (string line in BuildEdgeSummaryLines(groups))
            {
                sb.AppendLine(EscapeCsv(line));
            }

            File.WriteAllText(path, sb.ToString(), Encoding.UTF8);
            return path;
        }

        private static string WriteXlsx(Database db, IList<BoardGroup> groups, string assemblyNumber)
        {
            string drawingPath = db.Filename;
            string directory = string.IsNullOrWhiteSpace(drawingPath)
                ? Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory)
                : Path.GetDirectoryName(drawingPath);

            string name = string.IsNullOrWhiteSpace(drawingPath)
                ? "BoardSort"
                : Path.GetFileNameWithoutExtension(drawingPath);

            string path = Path.Combine(directory, name + "_boards.xlsx");
            path = GetWritableOutputPath(path);
            Dictionary<string, int> coatingNumbers = BuildCoatingNumbers(groups);

            using (ZipArchive archive = ZipFile.Open(path, ZipArchiveMode.Create))
            {
                AddZipText(archive, "[Content_Types].xml", XlsxContentTypes());
                AddZipText(archive, "_rels/.rels", XlsxRootRels());
                AddZipText(archive, "xl/workbook.xml", XlsxWorkbook());
                AddZipText(archive, "xl/_rels/workbook.xml.rels", XlsxWorkbookRels());
                AddZipText(archive, "xl/styles.xml", XlsxStyles());
                AddZipText(archive, "xl/worksheets/sheet1.xml", XlsxSheet(groups, coatingNumbers, assemblyNumber));
                AddZipText(archive, "xl/worksheets/_rels/sheet1.xml.rels", XlsxSheetRels());
                AddZipText(archive, "xl/drawings/drawing1.xml", XlsxDrawing(groups));
                AddZipText(archive, "xl/drawings/_rels/drawing1.xml.rels", XlsxDrawingRels(groups));

                for (int i = 0; i < groups.Count; i++)
                {
                    AddZipText(archive, "xl/media/sketch" + (i + 1).ToString(CultureInfo.InvariantCulture) + ".svg", BuildSketchSvg(groups[i], coatingNumbers));
                }
            }

            return path;
        }

        private static string GetWritableOutputPath(string path)
        {
            if (!File.Exists(path))
            {
                return path;
            }

            try
            {
                File.Delete(path);
                return path;
            }
            catch (IOException)
            {
            }
            catch (UnauthorizedAccessException)
            {
            }

            string directory = Path.GetDirectoryName(path);
            string name = Path.GetFileNameWithoutExtension(path);
            string extension = Path.GetExtension(path);
            string stamp = DateTime.Now.ToString("yyyyMMdd_HHmmss", CultureInfo.InvariantCulture);
            return Path.Combine(directory, name + "_" + stamp + extension);
        }

        private static void AddZipText(ZipArchive archive, string entryName, string text)
        {
            ZipArchiveEntry entry = archive.CreateEntry(entryName, CompressionLevel.Optimal);
            using (Stream stream = entry.Open())
            using (var writer = new StreamWriter(stream, new UTF8Encoding(false)))
            {
                writer.Write(text);
            }
        }

        private static string XlsxContentTypes()
        {
            return "<?xml version=\"1.0\" encoding=\"UTF-8\"?>"
                + "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">"
                + "<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>"
                + "<Default Extension=\"xml\" ContentType=\"application/xml\"/>"
                + "<Default Extension=\"svg\" ContentType=\"image/svg+xml\"/>"
                + "<Override PartName=\"/xl/workbook.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml\"/>"
                + "<Override PartName=\"/xl/worksheets/sheet1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/>"
                + "<Override PartName=\"/xl/drawings/drawing1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.drawing+xml\"/>"
                + "<Override PartName=\"/xl/styles.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml\"/>"
                + "</Types>";
        }

        private static string XlsxRootRels()
        {
            return "<?xml version=\"1.0\" encoding=\"UTF-8\"?>"
                + "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
                + "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"xl/workbook.xml\"/>"
                + "</Relationships>";
        }

        private static string XlsxWorkbook()
        {
            return "<?xml version=\"1.0\" encoding=\"UTF-8\"?>"
                + "<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">"
                + "<sheets><sheet name=\"Спецификация\" sheetId=\"1\" r:id=\"rId1\"/></sheets>"
                + "</workbook>";
        }

        private static string XlsxWorkbookRels()
        {
            return "<?xml version=\"1.0\" encoding=\"UTF-8\"?>"
                + "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
                + "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/sheet1.xml\"/>"
                + "<Relationship Id=\"rId2\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles\" Target=\"styles.xml\"/>"
                + "</Relationships>";
        }

        private static string XlsxStyles()
        {
            return "<?xml version=\"1.0\" encoding=\"UTF-8\"?>"
                + "<styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">"
                + "<fonts count=\"2\"><font><sz val=\"11\"/><name val=\"Calibri\"/></font><font><b/><sz val=\"11\"/><name val=\"Calibri\"/></font></fonts>"
                + "<fills count=\"2\"><fill><patternFill patternType=\"none\"/></fill><fill><patternFill patternType=\"gray125\"/></fill></fills>"
                + "<borders count=\"1\"><border><left/><right/><top/><bottom/><diagonal/></border></borders>"
                + "<cellStyleXfs count=\"1\"><xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\"/></cellStyleXfs>"
                + "<cellXfs count=\"3\"><xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\" applyAlignment=\"1\"><alignment horizontal=\"center\" vertical=\"center\"/></xf><xf numFmtId=\"0\" fontId=\"1\" fillId=\"0\" borderId=\"0\" xfId=\"0\" applyFont=\"1\" applyAlignment=\"1\"><alignment horizontal=\"center\" vertical=\"center\"/></xf><xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\" applyAlignment=\"1\"><alignment horizontal=\"center\" vertical=\"center\" wrapText=\"1\"/></xf></cellXfs>"
                + "<cellStyles count=\"1\"><cellStyle name=\"Normal\" xfId=\"0\" builtinId=\"0\"/></cellStyles>"
                + "</styleSheet>";
        }

        private static string XlsxSheet(IList<BoardGroup> groups, Dictionary<string, int> coatingNumbers, string assemblyNumber)
        {
            string[] headers =
            {
                "Сборка", "Деталь", "Наименование", "Материал", "Длина", "Ширина", "Толщина", "Кол-во",
                "P1", "P2", "L1", "L2", "W1", "W2", "Примечание", "Изображение"
            };

            var sb = new StringBuilder();
            sb.Append("<?xml version=\"1.0\" encoding=\"UTF-8\"?>");
            sb.Append("<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">");
            sb.Append("<cols>");
            for (int i = 1; i <= headers.Length; i++)
            {
                double width = i == 3 ? 20.0 : (i == 4 ? 24.0 : (i == 15 ? 30.0 : (i == 16 ? 46.0 : (i == 8 ? 9.0 : (i >= 9 && i <= 14 ? 9.0 : 13.0)))));
                sb.Append("<col min=\"").Append(i.ToString(CultureInfo.InvariantCulture))
                    .Append("\" max=\"").Append(i.ToString(CultureInfo.InvariantCulture))
                    .Append("\" width=\"").Append(width.ToString("0.#", CultureInfo.InvariantCulture))
                    .Append("\" customWidth=\"1\"/>");
            }
            sb.Append("</cols>");
            sb.Append("<sheetData>");

            sb.Append("<row r=\"1\">");
            for (int i = 0; i < headers.Length; i++)
            {
                AppendInlineCell(sb, i + 1, 1, headers[i], 1);
            }
            sb.Append("</row>");

            int row = 2;
            foreach (BoardGroup group in groups)
            {
                sb.Append("<row r=\"").Append(row.ToString(CultureInfo.InvariantCulture)).Append("\" ht=\"98\" customHeight=\"1\">");
                AppendInlineCell(sb, 1, row, assemblyNumber, 0);
                AppendNumberCell(sb, 2, row, group.Number);
                AppendInlineCell(sb, 3, row, DisplayPartName(group), 0);
                AppendInlineCell(sb, 4, row, DisplayMaterial(group), 0);
                AppendNumberCell(sb, 5, row, DisplayLength(group));
                AppendNumberCell(sb, 6, row, DisplayWidth(group));
                AppendNumberCell(sb, 7, row, group.ThicknessMm);
                AppendNumberCell(sb, 8, row, group.Quantity);
                AppendInlineCell(sb, 9, row, CoatingNumberText(Slot(group.Coatings, "P1"), coatingNumbers), 0);
                AppendInlineCell(sb, 10, row, CoatingNumberText(Slot(group.Coatings, "P2"), coatingNumbers), 0);
                AppendInlineCell(sb, 11, row, EdgeCoatingNumberText(Slot(group.Coatings, "L1"), coatingNumbers), 0);
                AppendInlineCell(sb, 12, row, EdgeCoatingNumberText(Slot(group.Coatings, "L2"), coatingNumbers), 0);
                AppendInlineCell(sb, 13, row, EdgeCoatingNumberText(Slot(group.Coatings, "W1"), coatingNumbers), 0);
                AppendInlineCell(sb, 14, row, EdgeCoatingNumberText(Slot(group.Coatings, "W2"), coatingNumbers), 0);
                AppendInlineCell(sb, 15, row, BuildGeneratedNote(group).Replace("; ", "\n"), 2);
                sb.Append("</row>");
                row++;
            }

            int legendStart = row + 2;
            sb.Append("<row r=\"").Append(legendStart.ToString(CultureInfo.InvariantCulture)).Append("\">");
            AppendInlineCell(sb, 1, legendStart, "Отделки", 1);
            sb.Append("</row>");
            row = legendStart + 1;
            foreach (KeyValuePair<string, int> pair in coatingNumbers.OrderBy(x => x.Value))
            {
                sb.Append("<row r=\"").Append(row.ToString(CultureInfo.InvariantCulture)).Append("\">");
                AppendNumberCell(sb, 1, row, pair.Value);
                AppendInlineCell(sb, 2, row, pair.Key, 0);
                sb.Append("</row>");
                row++;
            }

            row++;
            AppendSummarySection(sb, ref row, 1, "Итоги по материалам", BuildMaterialSummaryLines(groups));
            AppendSummarySection(sb, ref row, 1, "Итоги по кромке", BuildEdgeSummaryLines(groups));
            sb.Append("</sheetData>");

            sb.Append("<autoFilter ref=\"A1:P").Append(Math.Max(1, groups.Count + 1).ToString(CultureInfo.InvariantCulture)).Append("\"/>");
            if (groups.Count > 0)
            {
                sb.Append("<drawing r:id=\"rId1\"/>");
            }
            sb.Append("</worksheet>");
            return sb.ToString();
        }

        private static string XlsxSheetRels()
        {
            return "<?xml version=\"1.0\" encoding=\"UTF-8\"?>"
                + "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
                + "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing\" Target=\"../drawings/drawing1.xml\"/>"
                + "</Relationships>";
        }

        private static string XlsxDrawingRels(IList<BoardGroup> groups)
        {
            var sb = new StringBuilder();
            sb.Append("<?xml version=\"1.0\" encoding=\"UTF-8\"?>");
            sb.Append("<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">");
            for (int i = 0; i < groups.Count; i++)
            {
                int id = i + 1;
                sb.Append("<Relationship Id=\"rId").Append(id.ToString(CultureInfo.InvariantCulture))
                    .Append("\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/image\" Target=\"../media/sketch")
                    .Append(id.ToString(CultureInfo.InvariantCulture)).Append(".svg\"/>");
            }
            sb.Append("</Relationships>");
            return sb.ToString();
        }

        private static string XlsxDrawing(IList<BoardGroup> groups)
        {
            var sb = new StringBuilder();
            sb.Append("<?xml version=\"1.0\" encoding=\"UTF-8\"?>");
            sb.Append("<xdr:wsDr xmlns:xdr=\"http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing\" xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">");

            for (int i = 0; i < groups.Count; i++)
            {
                int rowIndex = i + 1;
                int imageId = i + 1;
                sb.Append("<xdr:oneCellAnchor>")
                    .Append("<xdr:from><xdr:col>15</xdr:col><xdr:colOff>15000</xdr:colOff><xdr:row>")
                    .Append(rowIndex.ToString(CultureInfo.InvariantCulture))
                    .Append("</xdr:row><xdr:rowOff>15000</xdr:rowOff></xdr:from>")
                    .Append("<xdr:ext cx=\"3200400\" cy=\"1238250\"/>")
                    .Append("<xdr:pic>")
                    .Append("<xdr:nvPicPr><xdr:cNvPr id=\"").Append(imageId.ToString(CultureInfo.InvariantCulture))
                    .Append("\" name=\"Sketch ").Append(imageId.ToString(CultureInfo.InvariantCulture))
                    .Append("\"/><xdr:cNvPicPr><a:picLocks noChangeAspect=\"1\"/></xdr:cNvPicPr></xdr:nvPicPr>")
                    .Append("<xdr:blipFill><a:blip r:embed=\"rId").Append(imageId.ToString(CultureInfo.InvariantCulture))
                    .Append("\"/><a:stretch><a:fillRect/></a:stretch></xdr:blipFill>")
                    .Append("<xdr:spPr><a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom></xdr:spPr>")
                    .Append("</xdr:pic><xdr:clientData/></xdr:oneCellAnchor>");
            }

            sb.Append("</xdr:wsDr>");
            return sb.ToString();
        }

        private static Dictionary<string, int> BuildCoatingNumbers(IList<BoardGroup> groups)
        {
            var result = new Dictionary<string, int>(StringComparer.CurrentCultureIgnoreCase);

            foreach (BoardGroup group in groups)
            {
                AddCoatingNumber(result, Slot(group.Coatings, "P1"));
                AddCoatingNumber(result, Slot(group.Coatings, "P2"));
                AddCoatingNumber(result, Slot(group.Coatings, "L1"));
                AddCoatingNumber(result, Slot(group.Coatings, "L2"));
                AddCoatingNumber(result, Slot(group.Coatings, "W1"));
                AddCoatingNumber(result, Slot(group.Coatings, "W2"));
                foreach (string item in EnumerateSketchCoatings(group))
                {
                    AddCoatingNumber(result, item);
                }
            }

            return result;
        }

        private static void AddCoatingNumber(Dictionary<string, int> numbers, string coating)
        {
            foreach (string item in SplitCoatings(coating))
            {
                if (!numbers.ContainsKey(item))
                {
                    numbers.Add(item, numbers.Count + 1);
                }
            }
        }

        private static string CoatingNumberText(string coating, Dictionary<string, int> numbers)
        {
            return string.Join("/", SplitCoatings(coating).Select(x => numbers.ContainsKey(x) ? numbers[x].ToString(CultureInfo.InvariantCulture) : string.Empty).Where(x => !string.IsNullOrWhiteSpace(x)));
        }

        private static string EdgeCoatingNumberText(string coating, Dictionary<string, int> numbers)
        {
            string first = SplitCoatings(coating).FirstOrDefault();
            if (string.IsNullOrWhiteSpace(first) || !numbers.ContainsKey(first))
            {
                return string.Empty;
            }

            return numbers[first].ToString(CultureInfo.InvariantCulture);
        }

        private static int EdgeCoatingNumber(string coating, Dictionary<string, int> numbers)
        {
            string text = EdgeCoatingNumberText(coating, numbers);
            int value;
            return int.TryParse(text, NumberStyles.Integer, CultureInfo.InvariantCulture, out value) ? value : 0;
        }

        private static IEnumerable<string> SplitCoatings(string coating)
        {
            if (string.IsNullOrWhiteSpace(coating))
            {
                yield break;
            }

            string[] parts = coating.Split(new[] { " / " }, StringSplitOptions.RemoveEmptyEntries);
            foreach (string part in parts)
            {
                string value = part.Trim();
                if (!string.IsNullOrWhiteSpace(value))
                {
                    yield return value;
                }
            }
        }

        private static string BuildTextSketch(BoardGroup group, Dictionary<string, int> coatingNumbers)
        {
            string l1 = EdgeCoatingNumberText(Slot(group.Coatings, "L1"), coatingNumbers);
            string l2 = EdgeCoatingNumberText(Slot(group.Coatings, "L2"), coatingNumbers);
            string w1 = EdgeCoatingNumberText(Slot(group.Coatings, "W1"), coatingNumbers);
            string w2 = EdgeCoatingNumberText(Slot(group.Coatings, "W2"), coatingNumbers);
            string p1 = CoatingNumberText(Slot(group.Coatings, "P1"), coatingNumbers);
            string p2 = CoatingNumberText(Slot(group.Coatings, "P2"), coatingNumbers);

            var parts = new List<string>();
            if (!string.IsNullOrWhiteSpace(l1)) parts.Add("верх=" + l1);
            if (!string.IsNullOrWhiteSpace(l2)) parts.Add("низ=" + l2);
            if (!string.IsNullOrWhiteSpace(w1)) parts.Add("право=" + w1);
            if (!string.IsNullOrWhiteSpace(w2)) parts.Add("лево=" + w2);
            if (!string.IsNullOrWhiteSpace(p1) || !string.IsNullOrWhiteSpace(p2))
            {
                parts.Add("пласть=" + string.Join("/", new[] { p1, p2 }.Where(x => !string.IsNullOrWhiteSpace(x))));
            }

            return parts.Count == 0 ? string.Empty : string.Join("; ", parts);
        }

        private static string BuildGeneratedNote(BoardGroup group)
        {
            List<string> items = BuildEdgeLengthItems(group);
            return items.Count == 0 ? string.Empty : string.Join("; ", items);
        }

        private static IEnumerable<string> EnumerateSketchCoatings(BoardGroup group)
        {
            if (group == null || group.Sketch == null)
            {
                yield break;
            }

            foreach (BoardSketchEdge edge in group.Sketch.Edges)
            {
                foreach (string coating in SplitCoatings(edge.Coating))
                {
                    yield return coating;
                }
            }
        }

        private static List<string> BuildEdgeLengthItems(BoardGroup group)
        {
            var totals = new Dictionary<string, double>(StringComparer.CurrentCultureIgnoreCase);
            foreach (KeyValuePair<string, double> pair in BuildEdgeLengthTotals(group))
            {
                string label = pair.Key + " (" + Format(group.ThicknessMm) + " мм)";
                totals[label] = pair.Value;
            }

            return totals
                .OrderByDescending(x => x.Value)
                .Select(x => x.Key + ": " + Format(x.Value) + " мм")
                .ToList();
        }

        private static Dictionary<string, double> BuildEdgeLengthTotals(BoardGroup group)
        {
            var result = new Dictionary<string, double>(StringComparer.CurrentCultureIgnoreCase);
            if (group == null || group.Sketch == null)
            {
                return result;
            }

            foreach (BoardSketchEdge edge in group.Sketch.Edges)
            {
                if (edge.StartIndex < 0 || edge.StartIndex >= group.Sketch.Points.Count || edge.EndIndex < 0 || edge.EndIndex >= group.Sketch.Points.Count)
                {
                    continue;
                }

                string coating = FirstCoating(edge.Coating);
                if (string.IsNullOrWhiteSpace(coating))
                {
                    continue;
                }

                BoardSketchPoint a = group.Sketch.Points[edge.StartIndex];
                BoardSketchPoint b = group.Sketch.Points[edge.EndIndex];
                double length = Math.Sqrt(Math.Pow(b.X - a.X, 2.0) + Math.Pow(b.Y - a.Y, 2.0));
                if (length <= 0.01)
                {
                    continue;
                }

                if (result.ContainsKey(coating))
                {
                    result[coating] += length;
                }
                else
                {
                    result.Add(coating, length);
                }
            }

            return result;
        }

        private static string FirstCoating(string coating)
        {
            return SplitCoatings(coating).FirstOrDefault() ?? string.Empty;
        }

        private static void PrintMaterialSummary(Editor ed, IList<BoardGroup> groups)
        {
            List<string> lines = BuildMaterialSummaryLines(groups);
            if (lines.Count == 0)
            {
                return;
            }

            ed.WriteMessage("\nИтоги по материалам:");
            foreach (string line in lines)
            {
                ed.WriteMessage("\n  {0}", line);
            }
        }

        private static void PrintEdgeSummary(Editor ed, IList<BoardGroup> groups)
        {
            List<string> lines = BuildEdgeSummaryLines(groups);
            if (lines.Count == 0)
            {
                return;
            }

            ed.WriteMessage("\nИтоги по кромке:");
            foreach (string line in lines)
            {
                ed.WriteMessage("\n  {0}", line);
            }
        }

        private static List<string> BuildMaterialSummaryLines(IList<BoardGroup> groups)
        {
            var square = new Dictionary<string, double>(StringComparer.CurrentCultureIgnoreCase);
            var linear = new Dictionary<string, double>(StringComparer.CurrentCultureIgnoreCase);
            var volume = new Dictionary<string, double>(StringComparer.CurrentCultureIgnoreCase);

            foreach (BoardGroup group in groups)
            {
                string material = DefaultMaterial(group);
                string kind = ClassifyGroupMaterialKind(group);
                double qty = group.Quantity;
                if (kind == "Погонный")
                {
                    AddToSummary(linear, material, group.LengthMm * qty / 1000.0);
                }
                else if (kind == "Объемный")
                {
                    AddToSummary(volume, material, group.LengthMm * group.WidthMm * group.ThicknessMm * qty / 1000000000.0);
                }
                else
                {
                    AddToSummary(square, material, group.LengthMm * group.WidthMm * qty / 1000000.0);
                }
            }

            var result = new List<string>();
            result.AddRange(square.OrderBy(x => x.Key).Select(x => x.Key + ": " + Format(x.Value) + " м2"));
            result.AddRange(linear.OrderBy(x => x.Key).Select(x => x.Key + ": " + Format(x.Value) + " пог. м"));
            result.AddRange(volume.OrderBy(x => x.Key).Select(x => x.Key + ": " + Format(x.Value) + " м3"));
            return result;
        }

        private static List<string> BuildEdgeSummaryLines(IList<BoardGroup> groups)
        {
            var totals = new Dictionary<string, double>(StringComparer.CurrentCultureIgnoreCase);
            foreach (BoardGroup group in groups)
            {
                foreach (KeyValuePair<string, double> pair in BuildEdgeLengthTotals(group))
                {
                    string key = pair.Key + " (" + Format(group.ThicknessMm) + " мм)";
                    AddToSummary(totals, key, pair.Value * group.Quantity / 1000.0);
                }
            }

            return totals.OrderBy(x => x.Key).Select(x => x.Key + ": " + Format(x.Value) + " пог. м").ToList();
        }

        private static void AddToSummary(IDictionary<string, double> totals, string key, double value)
        {
            if (string.IsNullOrWhiteSpace(key) || value <= 0.0)
            {
                return;
            }

            if (totals.ContainsKey(key))
            {
                totals[key] += value;
            }
            else
            {
                totals.Add(key, value);
            }
        }

        private static string ClassifyGroupMaterialKind(BoardGroup group)
        {
            double dimsOverThickness = Math.Max(group.LengthMm, group.WidthMm) / Math.Max(group.ThicknessMm, 1.0);
            if (group.WidthMm <= group.ThicknessMm * 2.0 || group.LengthMm <= group.ThicknessMm * 2.0)
            {
                return "Погонный";
            }

            if (dimsOverThickness < 4.0)
            {
                return "Объемный";
            }

            return "Площадной";
        }

        private static void AppendSummarySection(StringBuilder sb, ref int row, int column, string title, IList<string> lines)
        {
            if (lines == null || lines.Count == 0)
            {
                return;
            }

            sb.Append("<row r=\"").Append(row.ToString(CultureInfo.InvariantCulture)).Append("\">");
            AppendInlineCell(sb, column, row, title, 1);
            sb.Append("</row>");
            row++;

            foreach (string line in lines)
            {
                sb.Append("<row r=\"").Append(row.ToString(CultureInfo.InvariantCulture)).Append("\">");
                AppendInlineCell(sb, column, row, line, 0);
                sb.Append("</row>");
                row++;
            }
        }

        private static string BuildSketchSvg(BoardGroup group, Dictionary<string, int> coatingNumbers)
        {
            if (group.Sketch != null && group.Sketch.HasGeometry)
            {
                return BuildRealContourSketchSvg(group, coatingNumbers);
            }

            string l1 = EdgeCoatingNumberText(Slot(group.Coatings, "L1"), coatingNumbers);
            string l2 = EdgeCoatingNumberText(Slot(group.Coatings, "L2"), coatingNumbers);
            string w1 = EdgeCoatingNumberText(Slot(group.Coatings, "W1"), coatingNumbers);
            string w2 = EdgeCoatingNumberText(Slot(group.Coatings, "W2"), coatingNumbers);
            string p1 = CoatingNumberText(Slot(group.Coatings, "P1"), coatingNumbers);
            string p2 = CoatingNumberText(Slot(group.Coatings, "P2"), coatingNumbers);

            string center = DisplayLength(group).ToString("0.#", CultureInfo.InvariantCulture) + " x " + DisplayWidth(group).ToString("0.#", CultureInfo.InvariantCulture);
            string faceText = string.Join(" / ", new[] { p1, p2 }.Where(x => !string.IsNullOrWhiteSpace(x)));

            const double canvasW = 336.0;
            const double canvasH = 130.0;
            const double marginX = 24.0;
            const double marginTop = 42.0;
            const double marginBottom = 8.0;
            double length = Math.Max(1.0, DisplayLength(group));
            double width = Math.Max(1.0, DisplayWidth(group));
            double scaleX = (canvasW - marginX * 2.0) / length;
            double scaleY = (canvasH - marginTop - marginBottom) / width;
            double scale = Math.Min(scaleX, scaleY);
            double rectW = Math.Max(64.0, length * scale);
            double rectH = Math.Max(18.0, width * scale);
            double x = (canvasW - rectW) / 2.0;
            double y = marginTop + ((canvasH - marginTop - marginBottom) - rectH) / 2.0;
            double right = x + rectW;
            double bottom = y + rectH;

            var sb = new StringBuilder();
            sb.Append("<?xml version=\"1.0\" encoding=\"UTF-8\"?>");
            sb.Append("<svg xmlns=\"http://www.w3.org/2000/svg\" width=\"336\" height=\"130\" viewBox=\"0 0 336 130\">");
            sb.Append("<rect width=\"336\" height=\"130\" fill=\"#ffffff\"/>");
            sb.Append("<rect x=\"").Append(SvgNumber(x)).Append("\" y=\"").Append(SvgNumber(y))
                .Append("\" width=\"").Append(SvgNumber(rectW)).Append("\" height=\"").Append(SvgNumber(rectH))
                .Append("\" fill=\"#f8fafc\" stroke=\"#334155\" stroke-width=\"2\"/>");

            HashSet<int> labeledEdgeIndexes = FindLongestRectLabelEdges(length, width, l1, l2, w1, w2);
            AppendEdgeLine(sb, x, y, right, y, l1, (x + right) / 2.0, y - 5.0, labeledEdgeIndexes.Contains(0));
            AppendEdgeLine(sb, x, bottom, right, bottom, l2, (x + right) / 2.0, bottom + 5.0, labeledEdgeIndexes.Contains(1));
            AppendEdgeLine(sb, right, y, right, bottom, w1, right + 5.0, (y + bottom) / 2.0, labeledEdgeIndexes.Contains(2));
            AppendEdgeLine(sb, x, y, x, bottom, w2, x - 5.0, (y + bottom) / 2.0, labeledEdgeIndexes.Contains(3));

            sb.Append("<text x=\"168\" y=\"16\" font-family=\"Segoe UI, Arial, sans-serif\" font-size=\"18\" font-weight=\"700\" text-anchor=\"middle\" dominant-baseline=\"middle\" fill=\"#111827\">")
                .Append(EscapeXml(center)).Append("</text>");
            sb.Append("<text x=\"168\" y=\"31\" font-family=\"Segoe UI, Arial, sans-serif\" font-size=\"12\" text-anchor=\"middle\" fill=\"#64748b\">")
                .Append(EscapeXml(group.ThicknessMm.ToString("0.#", CultureInfo.InvariantCulture) + " мм"))
                .Append("</text>");
            if (!string.IsNullOrWhiteSpace(faceText))
            {
                sb.Append("<text x=\"168\" y=\"65\" font-family=\"Segoe UI, Arial, sans-serif\" font-size=\"11\" text-anchor=\"middle\" dominant-baseline=\"middle\" fill=\"#475569\">")
                    .Append(EscapeXml("Пласть: " + faceText))
                    .Append("</text>");
            }
            sb.Append("</svg>");
            return sb.ToString();
        }

        private static string BuildRealContourSketchSvg(BoardGroup group, Dictionary<string, int> coatingNumbers)
        {
            string p1 = CoatingNumberText(Slot(group.Coatings, "P1"), coatingNumbers);
            string p2 = CoatingNumberText(Slot(group.Coatings, "P2"), coatingNumbers);
            string center = DisplayLength(group).ToString("0.#", CultureInfo.InvariantCulture) + " x " + DisplayWidth(group).ToString("0.#", CultureInfo.InvariantCulture);
            string faceText = string.Join(" / ", new[] { p1, p2 }.Where(x => !string.IsNullOrWhiteSpace(x)));

            const double canvasW = 336.0;
            const double canvasH = 130.0;
            const double marginX = 24.0;
            const double marginTop = 42.0;
            const double marginBottom = 8.0;
            List<BoardSketchPoint> normalized = NormalizeSketchOrientation(group.Sketch);
            double minX = normalized.Min(x => x.X);
            double maxX = normalized.Max(x => x.X);
            double minY = normalized.Min(x => x.Y);
            double maxY = normalized.Max(x => x.Y);
            double spanX = Math.Max(1.0, maxX - minX);
            double spanY = Math.Max(1.0, maxY - minY);
            double scale = Math.Min((canvasW - marginX * 2.0) / spanX, (canvasH - marginTop - marginBottom) / spanY);
            double drawingW = spanX * scale;
            double drawingH = spanY * scale;
            double offsetX = (canvasW - drawingW) / 2.0;
            double offsetY = marginTop + ((canvasH - marginTop - marginBottom) - drawingH) / 2.0;

            var projected = normalized
                .Select(point => new BoardSketchPoint
                {
                    X = offsetX + (point.X - minX) * scale,
                    Y = offsetY + (maxY - point.Y) * scale
                })
                .ToList();

            var sb = new StringBuilder();
            sb.Append("<?xml version=\"1.0\" encoding=\"UTF-8\"?>");
            sb.Append("<svg xmlns=\"http://www.w3.org/2000/svg\" width=\"336\" height=\"130\" viewBox=\"0 0 336 130\">");
            sb.Append("<rect width=\"336\" height=\"130\" fill=\"#ffffff\"/>");

            HashSet<int> labeledEdgeIndexes = FindLongestLabeledEdges(group.Sketch, projected, coatingNumbers);
            int edgeIndex = 0;

            foreach (BoardSketchEdge edge in group.Sketch.Edges)
            {
                if (edge.StartIndex < 0 || edge.StartIndex >= projected.Count || edge.EndIndex < 0 || edge.EndIndex >= projected.Count)
                {
                    continue;
                }

                BoardSketchPoint a = projected[edge.StartIndex];
                BoardSketchPoint b = projected[edge.EndIndex];
                string number = EdgeCoatingNumberText(edge.Coating, coatingNumbers);
                if (edge.IsArc && edge.ArcRadius > 0.0)
                {
                    AppendEdgeArc(sb, a.X, a.Y, b.X, b.Y, edge.ArcRadius * scale, edge.ArcLarge, !edge.ArcSweep, number, (a.X + b.X) / 2.0, Math.Min(a.Y, b.Y) - 5.0, labeledEdgeIndexes.Contains(edgeIndex) && !string.IsNullOrWhiteSpace(number));
                }
                else
                {
                    AppendEdgeLine(sb, a.X, a.Y, b.X, b.Y, number, (a.X + b.X) / 2.0, (a.Y + b.Y) / 2.0, labeledEdgeIndexes.Contains(edgeIndex) && !string.IsNullOrWhiteSpace(number));
                }

                edgeIndex++;
            }

            sb.Append("<text x=\"168\" y=\"16\" font-family=\"Segoe UI, Arial, sans-serif\" font-size=\"18\" font-weight=\"700\" text-anchor=\"middle\" dominant-baseline=\"middle\" fill=\"#111827\">")
                .Append(EscapeXml(center)).Append("</text>");
            sb.Append("<text x=\"168\" y=\"31\" font-family=\"Segoe UI, Arial, sans-serif\" font-size=\"12\" text-anchor=\"middle\" fill=\"#64748b\">")
                .Append(EscapeXml(group.ThicknessMm.ToString("0.#", CultureInfo.InvariantCulture) + " мм"))
                .Append("</text>");
            if (!string.IsNullOrWhiteSpace(faceText))
            {
                sb.Append("<text x=\"168\" y=\"65\" font-family=\"Segoe UI, Arial, sans-serif\" font-size=\"11\" text-anchor=\"middle\" dominant-baseline=\"middle\" fill=\"#475569\">")
                    .Append(EscapeXml("Пласть: " + faceText))
                    .Append("</text>");
            }
            sb.Append("</svg>");
            return sb.ToString();
        }

        private static HashSet<int> FindLongestRectLabelEdges(double length, double width, string l1, string l2, string w1, string w2)
        {
            var bestIndexes = new Dictionary<string, int>(StringComparer.CurrentCultureIgnoreCase);
            var bestLengths = new Dictionary<string, double>(StringComparer.CurrentCultureIgnoreCase);
            var bestRanks = new Dictionary<string, double>(StringComparer.CurrentCultureIgnoreCase);

            TrySetBestRectLabelEdge(bestIndexes, bestLengths, bestRanks, l1, 0, length, 0.0);
            TrySetBestRectLabelEdge(bestIndexes, bestLengths, bestRanks, l2, 1, length, 1.0);
            TrySetBestRectLabelEdge(bestIndexes, bestLengths, bestRanks, w1, 2, width, 2.0);
            TrySetBestRectLabelEdge(bestIndexes, bestLengths, bestRanks, w2, 3, width, 3.0);

            return new HashSet<int>(bestIndexes.Values);
        }

        private static void TrySetBestRectLabelEdge(
            IDictionary<string, int> bestIndexes,
            IDictionary<string, double> bestLengths,
            IDictionary<string, double> bestRanks,
            string coatingNumber,
            int edgeIndex,
            double edgeLength,
            double rank)
        {
            if (string.IsNullOrWhiteSpace(coatingNumber))
            {
                return;
            }

            double knownLength;
            double knownRank;
            bool hasKnownLength = bestLengths.TryGetValue(coatingNumber, out knownLength);
            bool hasKnownRank = bestRanks.TryGetValue(coatingNumber, out knownRank);
            if (!hasKnownLength
                || edgeLength > knownLength + 0.01
                || (Math.Abs(edgeLength - knownLength) < 0.01 && (!hasKnownRank || rank < knownRank)))
            {
                bestIndexes[coatingNumber] = edgeIndex;
                bestLengths[coatingNumber] = edgeLength;
                bestRanks[coatingNumber] = rank;
            }
        }

        private static HashSet<int> FindLongestLabeledEdges(BoardSketch sketch, IList<BoardSketchPoint> points, Dictionary<string, int> coatingNumbers)
        {
            var result = new HashSet<int>();
            if (sketch == null || points == null)
            {
                return result;
            }

            var bestIndexes = new Dictionary<string, int>(StringComparer.CurrentCultureIgnoreCase);
            var bestLengths = new Dictionary<string, double>(StringComparer.CurrentCultureIgnoreCase);
            var bestTops = new Dictionary<string, double>(StringComparer.CurrentCultureIgnoreCase);

            for (int i = 0; i < sketch.Edges.Count; i++)
            {
                BoardSketchEdge edge = sketch.Edges[i];
                string coatingNumber = EdgeCoatingNumberText(edge.Coating, coatingNumbers);
                if (string.IsNullOrWhiteSpace(coatingNumber)
                    || edge.StartIndex < 0 || edge.StartIndex >= points.Count
                    || edge.EndIndex < 0 || edge.EndIndex >= points.Count)
                {
                    continue;
                }

                BoardSketchPoint a = points[edge.StartIndex];
                BoardSketchPoint b = points[edge.EndIndex];
                double length = Math.Sqrt(Math.Pow(b.X - a.X, 2.0) + Math.Pow(b.Y - a.Y, 2.0));
                double top = Math.Min(a.Y, b.Y);
                double knownLength;
                double knownTop;
                if (!bestLengths.TryGetValue(coatingNumber, out knownLength)
                    || length > knownLength + 0.01
                    || (Math.Abs(length - knownLength) < 0.01 && (!bestTops.TryGetValue(coatingNumber, out knownTop) || top < knownTop)))
                {
                    bestIndexes[coatingNumber] = i;
                    bestLengths[coatingNumber] = length;
                    bestTops[coatingNumber] = top;
                }
            }

            foreach (int index in bestIndexes.Values)
            {
                result.Add(index);
            }

            return result;
        }

        private static List<BoardSketchPoint> NormalizeSketchOrientation(BoardSketch sketch)
        {
            double angle = FindDominantSketchAngle(sketch);
            List<BoardSketchPoint> rotated = RotateSketchPoints(sketch.Points, -angle);
            double spanX = rotated.Max(x => x.X) - rotated.Min(x => x.X);
            double spanY = rotated.Max(x => x.Y) - rotated.Min(x => x.Y);
            if (spanY > spanX)
            {
                rotated = RotateSketchPoints(rotated, -Math.PI / 2.0);
            }

            return rotated;
        }

        private static double FindDominantSketchAngle(BoardSketch sketch)
        {
            double bestLength = -1.0;
            double bestAngle = 0.0;

            foreach (BoardSketchEdge edge in sketch.Edges)
            {
                if (edge.StartIndex < 0 || edge.StartIndex >= sketch.Points.Count || edge.EndIndex < 0 || edge.EndIndex >= sketch.Points.Count)
                {
                    continue;
                }

                BoardSketchPoint a = sketch.Points[edge.StartIndex];
                BoardSketchPoint b = sketch.Points[edge.EndIndex];
                double dx = b.X - a.X;
                double dy = b.Y - a.Y;
                double length = Math.Sqrt(dx * dx + dy * dy);
                if (length > bestLength)
                {
                    bestLength = length;
                    bestAngle = Math.Atan2(dy, dx);
                }
            }

            return bestAngle;
        }

        private static List<BoardSketchPoint> RotateSketchPoints(IEnumerable<BoardSketchPoint> points, double angle)
        {
            double c = Math.Cos(angle);
            double s = Math.Sin(angle);
            return points
                .Select(point => new BoardSketchPoint
                {
                    X = point.X * c - point.Y * s,
                    Y = point.X * s + point.Y * c
                })
                .ToList();
        }

        private static void AppendEdgeLine(StringBuilder sb, double x1, double y1, double x2, double y2, string coatingNumber, double textX, double textY)
        {
            AppendEdgeLine(sb, x1, y1, x2, y2, coatingNumber, textX, textY, true);
        }

        private static void AppendEdgeLine(StringBuilder sb, double x1, double y1, double x2, double y2, string coatingNumber, double textX, double textY, bool showLabel)
        {
            if (string.IsNullOrWhiteSpace(coatingNumber))
            {
                sb.Append("<line x1=\"").Append(SvgNumber(x1)).Append("\" y1=\"").Append(SvgNumber(y1))
                    .Append("\" x2=\"").Append(SvgNumber(x2)).Append("\" y2=\"").Append(SvgNumber(y2))
                    .Append("\" stroke=\"#94a3b8\" stroke-width=\"1\"/>");
                return;
            }

            string color = ColorForCoatingNumberText(coatingNumber);
            sb.Append("<line x1=\"").Append(SvgNumber(x1)).Append("\" y1=\"").Append(SvgNumber(y1))
                .Append("\" x2=\"").Append(SvgNumber(x2)).Append("\" y2=\"").Append(SvgNumber(y2))
                .Append("\" stroke=\"").Append(color).Append("\" stroke-width=\"4\" stroke-linecap=\"round\"/>");
            if (!showLabel)
            {
                return;
            }

            AppendEdgeMarker(sb, coatingNumber, color, textX, textY);
        }

        private static void AppendEdgeArc(StringBuilder sb, double x1, double y1, double x2, double y2, double radius, bool largeArc, bool sweep, string coatingNumber, double textX, double textY, bool showLabel)
        {
            double r = Math.Max(1.0, radius);
            string arcPath = "M " + SvgNumber(x1) + " " + SvgNumber(y1)
                + " A " + SvgNumber(r) + " " + SvgNumber(r)
                + " 0 " + (largeArc ? "1" : "0")
                + " " + (sweep ? "1" : "0")
                + " " + SvgNumber(x2) + " " + SvgNumber(y2);

            if (string.IsNullOrWhiteSpace(coatingNumber))
            {
                sb.Append("<path d=\"").Append(arcPath)
                    .Append("\" fill=\"none\" stroke=\"#94a3b8\" stroke-width=\"1\"/>");
                return;
            }

            string color = ColorForCoatingNumberText(coatingNumber);
            sb.Append("<path d=\"").Append(arcPath)
                .Append("\" fill=\"none\" stroke=\"").Append(color)
                .Append("\" stroke-width=\"4\" stroke-linecap=\"round\"/>");
            if (!showLabel)
            {
                return;
            }

            AppendEdgeMarker(sb, coatingNumber, color, textX, textY);
        }

        private static void AppendEdgeMarker(StringBuilder sb, string coatingNumber, string color, double textX, double textY)
        {
            sb.Append("<circle cx=\"").Append(SvgNumber(textX)).Append("\" cy=\"").Append(SvgNumber(textY))
                .Append("\" r=\"10\" fill=\"").Append(color).Append("\"/>");
            sb.Append("<text x=\"").Append(SvgNumber(textX)).Append("\" y=\"").Append(SvgNumber(textY))
                .Append("\" dy=\"0.35em\" font-family=\"Segoe UI, Arial, sans-serif\" font-size=\"11\" font-weight=\"700\" text-anchor=\"middle\" fill=\"#ffffff\">")
                .Append(EscapeXml(coatingNumber)).Append("</text>");
        }

        private static string ColorForCoatingNumberText(string numberText)
        {
            int number;
            return int.TryParse(numberText, NumberStyles.Integer, CultureInfo.InvariantCulture, out number)
                ? ColorForCoatingNumber(number)
                : "#94a3b8";
        }

        private static string ColorForCoatingNumber(int number)
        {
            string[] colors =
            {
                "#dc2626", "#2563eb", "#16a34a", "#f59e0b", "#9333ea", "#0891b2",
                "#db2777", "#65a30d", "#ea580c", "#4f46e5", "#0f766e", "#be123c"
            };

            if (number <= 0)
            {
                return "#94a3b8";
            }

            return colors[(number - 1) % colors.Length];
        }

        private static string ShortCoatingLabel(string coating)
        {
            if (string.IsNullOrWhiteSpace(coating))
            {
                return string.Empty;
            }

            string value = coating.Trim();
            if (value.Length <= 14)
            {
                return value;
            }

            return value.Substring(0, 13) + ".";
        }

        private static string SvgNumber(double value)
        {
            return value.ToString("0.###", CultureInfo.InvariantCulture);
        }

        private static void AppendInlineCell(StringBuilder sb, int column, int row, string value, int style)
        {
            sb.Append("<c r=\"").Append(CellReference(column, row)).Append("\" t=\"inlineStr\"");
            if (style > 0)
            {
                sb.Append(" s=\"").Append(style.ToString(CultureInfo.InvariantCulture)).Append("\"");
            }
            sb.Append("><is><t>");
            sb.Append(EscapeXml(value ?? string.Empty));
            sb.Append("</t></is></c>");
        }

        private static void AppendNumberCell(StringBuilder sb, int column, int row, double value)
        {
            sb.Append("<c r=\"").Append(CellReference(column, row)).Append("\"><v>")
                .Append(value.ToString("0.###", CultureInfo.InvariantCulture))
                .Append("</v></c>");
        }

        private static string CellReference(int column, int row)
        {
            var letters = new StringBuilder();
            int value = column;
            while (value > 0)
            {
                value--;
                letters.Insert(0, (char)('A' + value % 26));
                value /= 26;
            }

            return letters.ToString() + row.ToString(CultureInfo.InvariantCulture);
        }

        private static string EscapeXml(string value)
        {
            if (string.IsNullOrEmpty(value))
            {
                return string.Empty;
            }

            return value.Replace("&", "&amp;").Replace("<", "&lt;").Replace(">", "&gt;").Replace("\"", "&quot;");
        }

        private static string EscapeCsv(string value)
        {
            if (value == null)
            {
                return string.Empty;
            }

            if (value.IndexOfAny(new[] { ';', '"', '\r', '\n' }) < 0)
            {
                return value;
            }

            return "\"" + value.Replace("\"", "\"\"") + "\"";
        }

        private static string Format(double value)
        {
            return value.ToString("0.0", CultureInfo.InvariantCulture);
        }

        private static double Round1(double value)
        {
            return Math.Round(value, 1, MidpointRounding.AwayFromZero);
        }

        private static string TrimForColumn(string value, int width)
        {
            if (string.IsNullOrEmpty(value) || value.Length <= width)
            {
                return value ?? string.Empty;
            }

            return value.Substring(0, width - 1) + ".";
        }
    }
}
