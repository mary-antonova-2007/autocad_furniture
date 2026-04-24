using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.BoundaryRepresentation;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;
using AutoCAD_BoardSorter.Geometry;
using AutoCAD_BoardSorter.Models;
using AcApplication = Autodesk.AutoCAD.ApplicationServices.Application;
using BrepEdge = Autodesk.AutoCAD.BoundaryRepresentation.Edge;
using BrepFace = Autodesk.AutoCAD.BoundaryRepresentation.Face;

namespace AutoCAD_BoardSorter.Ui
{
    internal sealed class SpecificationPaletteControl : UserControl
    {
        private readonly TextBlock statusText;
        private readonly TextBox assemblyNumberBox;
        private readonly TextBox partNumberBox;
        private readonly TextBox partNameBox;
        private readonly ComboBox partTypeBox;
        private readonly TextBox lengthBox;
        private readonly TextBox widthBox;
        private readonly CheckBox rotateBox;
        private readonly TextBox materialBox;
        private readonly TextBox noteBox;
        private readonly TextBox coatingBox;
        private readonly TextBlock coatingStatusText;
        private readonly TextBox coatedFacesBox;
        private readonly StackPanel coatingPanel;
        private List<ObjectId> selectedSolidIds = new List<ObjectId>();
        private List<FaceSelectionRef> selectedFaces = new List<FaceSelectionRef>();
        private bool isRefreshing;

        public SpecificationPaletteControl()
        {
            var root = new DockPanel
            {
                LastChildFill = true,
                Background = new SolidColorBrush(Color.FromRgb(42, 48, 56))
            };

            var buttons = new StackPanel
            {
                Orientation = Orientation.Horizontal,
                Margin = new Thickness(10, 10, 10, 6)
            };

            var refreshButton = new Button { Content = "Обновить", MinWidth = 92, Margin = new Thickness(0, 0, 8, 0) };
            refreshButton.Click += delegate { RefreshFromSelection(); };
            buttons.Children.Add(refreshButton);

            var applyButton = new Button { Content = "Записать", MinWidth = 92 };
            applyButton.Click += delegate { ApplyToSelection(); };
            buttons.Children.Add(applyButton);

            DockPanel.SetDock(buttons, Dock.Top);
            root.Children.Add(buttons);

            statusText = new TextBlock
            {
                Foreground = Brushes.Gainsboro,
                Margin = new Thickness(10, 0, 10, 10),
                TextWrapping = TextWrapping.Wrap
            };
            DockPanel.SetDock(statusText, Dock.Top);
            root.Children.Add(statusText);

            var scroll = new ScrollViewer
            {
                VerticalScrollBarVisibility = ScrollBarVisibility.Auto
            };
            root.Children.Add(scroll);

            var grid = new Grid
            {
                Margin = new Thickness(10, 0, 10, 12)
            };
            grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(132) });
            grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
            scroll.Content = grid;

            int row = 0;
            assemblyNumberBox = AddTextRow(grid, ref row, "Номер сборки");
            partNumberBox = AddTextRow(grid, ref row, "Номер детали");
            partNameBox = AddTextRow(grid, ref row, "Наименование");

            partTypeBox = new ComboBox { Margin = FieldMargin() };
            partTypeBox.Items.Add("Погонный");
            partTypeBox.Items.Add("Площадной");
            partTypeBox.Items.Add("Объёмный");
            AddRow(grid, ref row, "Тип детали", partTypeBox);

            lengthBox = AddTextRow(grid, ref row, "Длина");
            widthBox = AddTextRow(grid, ref row, "Ширина");

            rotateBox = new CheckBox
            {
                Content = "Да",
                Foreground = Brushes.White,
                Margin = FieldMargin()
            };
            AddRow(grid, ref row, "Поворот", rotateBox);

            materialBox = AddTextRow(grid, ref row, "Материал");
            noteBox = AddTextRow(grid, ref row, "Примечание");

            coatingPanel = new StackPanel
            {
                Margin = new Thickness(10, 8, 10, 14)
            };
            DockPanel.SetDock(coatingPanel, Dock.Bottom);
            root.Children.Add(coatingPanel);

            coatingPanel.Children.Add(new TextBlock
            {
                Text = "Облицовка",
                FontWeight = FontWeights.Bold,
                Foreground = Brushes.White,
                Margin = new Thickness(0, 0, 0, 6)
            });

            coatingStatusText = new TextBlock
            {
                Foreground = Brushes.Gainsboro,
                TextWrapping = TextWrapping.Wrap,
                Margin = new Thickness(0, 0, 0, 4)
            };
            coatingPanel.Children.Add(coatingStatusText);

            coatingBox = new TextBox { Margin = new Thickness(0, 0, 0, 0) };
            coatingPanel.Children.Add(coatingBox);

            coatingPanel.Children.Add(new TextBlock
            {
                Text = "Облицованные грани выбранных тел",
                FontWeight = FontWeights.Bold,
                Foreground = Brushes.White,
                Margin = new Thickness(0, 10, 0, 4)
            });

            coatedFacesBox = new TextBox
            {
                MinHeight = 72,
                MaxHeight = 160,
                IsReadOnly = true,
                TextWrapping = TextWrapping.NoWrap,
                AcceptsReturn = true,
                VerticalScrollBarVisibility = ScrollBarVisibility.Auto,
                HorizontalScrollBarVisibility = ScrollBarVisibility.Auto,
                FontFamily = new FontFamily("Consolas"),
                Margin = new Thickness(0, 0, 0, 0)
            };
            coatingPanel.Children.Add(coatedFacesBox);

            SetFieldsEnabled(false);
            Content = root;
        }

        public void RefreshFromSelection()
        {
            if (isRefreshing)
            {
                return;
            }

            isRefreshing = true;
            try
            {
                Document doc = AcApplication.DocumentManager.MdiActiveDocument;
                if (doc == null)
                {
                    ClearSelectionState();
                    SetStatus("Нет активного документа.");
                    SetFieldsEnabled(false);
                    return;
                }

                PaletteDebugLogger.Info(doc.Database, "RefreshFromSelection: active document ok");
                Editor editor = doc.Editor;
                PaletteDebugLogger.Info(doc.Database, "RefreshFromSelection: SelectImplied start");
                PromptSelectionResult selection = editor.SelectImplied();
                PaletteDebugLogger.Info(doc.Database, "RefreshFromSelection: SelectImplied status=" + selection.Status);
                if (selection.Status != PromptStatus.OK)
                {
                    ClearSelectionState();
                    SetStatus("Выбери одно или несколько 3D-тел.");
                    RefreshFaceBlock(doc);
                    SetFieldsEnabled(false);
                    return;
                }

                ReadSelection(doc.Database, selection.Value, out selectedSolidIds, out selectedFaces);
                PaletteDebugLogger.Info(doc.Database, "RefreshFromSelection: solids="
                    + selectedSolidIds.Count.ToString(CultureInfo.InvariantCulture)
                    + " faces=" + selectedFaces.Count.ToString(CultureInfo.InvariantCulture));

                if (selectedSolidIds.Count == 0)
                {
                    ClearSelectionState();
                    SetStatus("В выделении нет 3D-тел.");
                    RefreshFaceBlock(doc);
                    SetFieldsEnabled(false);
                    return;
                }

                var values = new List<SpecificationData>();
                using (Transaction tr = doc.Database.TransactionManager.StartTransaction())
                {
                    PaletteDebugLogger.Info(doc.Database, "RefreshFromSelection: read specification start");
                    foreach (ObjectId id in selectedSolidIds)
                    {
                        PaletteDebugLogger.Info(doc.Database, "RefreshFromSelection: open solid " + id.Handle);
                        var solid = tr.GetObject(id, OpenMode.ForRead, false) as Solid3d;
                        if (solid == null || solid.IsErased)
                        {
                            continue;
                        }

                        SpecificationData data;
                        if (!SpecificationStorage.TryRead(solid, tr, out data))
                        {
                            data = new SpecificationData
                            {
                                PartName = "Деталь",
                                PartType = "Площадной",
                                RotateLengthWidth = false
                            };
                        }

                        values.Add(data);
                    }

                    tr.Commit();
                }
                PaletteDebugLogger.Info(doc.Database, "RefreshFromSelection: read specification finish values="
                    + values.Count.ToString(CultureInfo.InvariantCulture));

                if (values.Count == 0)
                {
                    ClearSelectionState();
                    SetStatus("Не удалось прочитать выбранные тела.");
                    SetFieldsEnabled(false);
                    return;
                }

                FillMerged(values);
                SetSelectionStatus();
                RefreshFaceBlock(doc);
                SetFieldsEnabled(true);
            }
            catch (System.Exception ex)
            {
                Document doc = AcApplication.DocumentManager.MdiActiveDocument;
                Database db = doc != null ? doc.Database : null;
                PaletteDebugLogger.Error(db, "RefreshFromSelection failed", ex);
                ClearSelectionState();
                SetStatus("Ошибка чтения выделения. Смотри palette log рядом с DWG.");
                SetFieldsEnabled(false);
            }
            finally
            {
                isRefreshing = false;
            }
        }

        private static List<FaceSelectionRef> BuildVisibleFacesFromCurrentView(Document doc, IList<ObjectId> solidIds)
        {
            var result = new List<FaceSelectionRef>();
            if (doc == null || solidIds == null || solidIds.Count == 0)
            {
                return result;
            }

            Vector3d viewDirection;
            try
            {
                viewDirection = doc.Editor.GetCurrentView().ViewDirection.GetNormal();
            }
            catch
            {
                viewDirection = Vector3d.ZAxis;
            }

            using (Transaction tr = doc.Database.TransactionManager.StartTransaction())
            {
                foreach (ObjectId id in solidIds)
                {
                    var solid = tr.GetObject(id, OpenMode.ForRead, false) as Solid3d;
                    if (solid == null || solid.IsErased)
                    {
                        continue;
                    }

                    string faceKey;
                    if (TryFindVisiblePlanarFaceKey(solid, viewDirection, out faceKey))
                    {
                        result.Add(new FaceSelectionRef(id, faceKey));
                    }
                }

                tr.Commit();
            }

            return result;
        }

        private static bool TryFindVisiblePlanarFaceKey(Solid3d solid, Vector3d viewDirection, out string faceKey)
        {
            faceKey = null;

            try
            {
                using (var brep = new Brep(solid))
                {
                    double bestScore = double.NegativeInfinity;
                    foreach (BrepFace face in brep.Faces)
                    {
                        Vector3d normal;
                        if (!TryGetPlanarFaceNormal(face, out normal))
                        {
                            continue;
                        }

                        double alignment = Math.Abs(normal.GetNormal().DotProduct(viewDirection));
                        if (alignment < 0.75)
                        {
                            continue;
                        }

                        BoundBlock3d bounds = face.BoundBlock;
                        Point3d min = bounds.GetMinimumPoint();
                        Point3d max = bounds.GetMaximumPoint();
                        Point3d center = new Point3d(
                            (min.X + max.X) / 2.0,
                            (min.Y + max.Y) / 2.0,
                            (min.Z + max.Z) / 2.0);

                        double score = alignment * 1000000.0 + center.GetAsVector().DotProduct(viewDirection);
                        string candidateKey;
                        if (score > bestScore && FaceKeyBuilder.TryBuild(face, out candidateKey))
                        {
                            bestScore = score;
                            faceKey = candidateKey;
                        }
                    }
                }
            }
            catch
            {
                return false;
            }

            return !string.IsNullOrWhiteSpace(faceKey);
        }

        private static bool TryGetPlanarFaceNormal(BrepFace face, out Vector3d normal)
        {
            normal = Vector3d.ZAxis;

            try
            {
                PlanarEntity planar = face.Surface as PlanarEntity;
                if (planar == null)
                {
                    return false;
                }

                object coefficients = planar.Coefficients;
                normal = new Vector3d(
                    GetDoubleMember(coefficients, "A", "a"),
                    GetDoubleMember(coefficients, "B", "b"),
                    GetDoubleMember(coefficients, "C", "c"));

                return normal.Length > 1e-9;
            }
            catch
            {
                return false;
            }
        }

        private static double GetDoubleMember(object instance, params string[] names)
        {
            if (instance == null)
            {
                return 0.0;
            }

            Type type = instance.GetType();
            for (int i = 0; i < names.Length; i++)
            {
                var property = type.GetProperty(names[i]);
                if (property != null)
                {
                    return Convert.ToDouble(property.GetValue(instance, null), CultureInfo.InvariantCulture);
                }

                var field = type.GetField(names[i]);
                if (field != null)
                {
                    return Convert.ToDouble(field.GetValue(instance), CultureInfo.InvariantCulture);
                }
            }

            return 0.0;
        }

        private void RefreshSpecificationFields(Document doc)
        {
            var values = new List<SpecificationData>();
            using (Transaction tr = doc.Database.TransactionManager.StartTransaction())
            {
                foreach (ObjectId id in selectedSolidIds)
                {
                    var solid = tr.GetObject(id, OpenMode.ForRead, false) as Solid3d;
                    if (solid == null || solid.IsErased)
                    {
                        continue;
                    }

                    SpecificationData data;
                    if (!SpecificationStorage.TryRead(solid, tr, out data))
                    {
                        data = new SpecificationData
                        {
                            PartName = "Деталь",
                            PartType = "Площадной",
                            RotateLengthWidth = false
                        };
                    }

                    values.Add(data);
                }

                tr.Commit();
            }

            if (values.Count == 0)
            {
                ClearFields();
                SetStatus("Не удалось прочитать выбранные тела.");
                return;
            }

            FillMerged(values);
            SetSelectionStatus();
        }

        private void ApplyToSelection()
        {
            Document doc = AcApplication.DocumentManager.MdiActiveDocument;
            if (doc == null || selectedSolidIds.Count == 0)
            {
                SetStatus("Нет выбранных 3D-тел.");
                return;
            }

            double length = 0.0;
            double width = 0.0;
            bool hasLength = !IsMergedValue(lengthBox.Text);
            bool hasWidth = !IsMergedValue(widthBox.Text);
            if ((hasLength && !TryParseLength(lengthBox.Text, out length))
                || (hasWidth && !TryParseLength(widthBox.Text, out width)))
            {
                SetStatus("Длина и ширина должны быть числами.");
                return;
            }

            using (doc.LockDocument())
            using (Transaction tr = doc.Database.TransactionManager.StartTransaction())
            {
                int written = 0;
                foreach (ObjectId id in selectedSolidIds)
                {
                    var solid = tr.GetObject(id, OpenMode.ForWrite, false) as Solid3d;
                    if (solid == null || solid.IsErased)
                    {
                        continue;
                    }

                    SpecificationData existing;
                    if (!SpecificationStorage.TryRead(solid, tr, out existing))
                    {
                        existing = new SpecificationData
                        {
                            PartName = "Деталь",
                            PartType = "Площадной"
                        };
                    }

                    var data = new SpecificationData
                    {
                        AssemblyNumber = FieldValue(assemblyNumberBox.Text, existing.AssemblyNumber),
                        PartNumber = FieldValue(partNumberBox.Text, existing.PartNumber),
                        PartName = FirstNonBlank(FieldValue(partNameBox.Text, existing.PartName), "Деталь"),
                        PartType = FirstNonBlank(FieldValue(partTypeBox.SelectedItem as string, existing.PartType), "Площадной"),
                        LengthMm = hasLength ? length : existing.LengthMm,
                        WidthMm = hasWidth ? width : existing.WidthMm,
                        RotateLengthWidth = rotateBox.IsChecked.HasValue ? rotateBox.IsChecked == true : existing.RotateLengthWidth,
                        Material = FieldValue(materialBox.Text, existing.Material),
                        Note = FieldValue(noteBox.Text, existing.Note)
                    };

                    SpecificationStorage.Write(solid, data, tr);
                    written++;
                }

                tr.Commit();
                SetStatus("Записано в тел: " + written.ToString(CultureInfo.InvariantCulture));
            }

            ApplyCoatingToFaces(doc);
        }

        private void FillMerged(IList<SpecificationData> values)
        {
            assemblyNumberBox.Text = Merge(values.Select(x => x.AssemblyNumber));
            partNumberBox.Text = Merge(values.Select(x => x.PartNumber));
            partNameBox.Text = Merge(values.Select(x => x.PartName));
            SetComboMerged(partTypeBox, values.Select(x => x.PartType));
            lengthBox.Text = Merge(values.Select(x => Format(x.LengthMm)));
            widthBox.Text = Merge(values.Select(x => Format(x.WidthMm)));
            SetCheckMerged(values.Select(x => x.RotateLengthWidth));
            materialBox.Text = Merge(values.Select(x => x.Material));
            noteBox.Text = Merge(values.Select(x => x.Note));
        }

        private void RefreshFaceBlock(Document doc)
        {
            bool hasFaces = selectedFaces.Count > 0;
            coatingPanel.IsEnabled = true;
            coatingBox.IsEnabled = hasFaces;
            coatingStatusText.Text = hasFaces
                ? "Выбрано граней: " + selectedFaces.Count.ToString(CultureInfo.InvariantCulture)
                : "Грани не выбраны. Выбери грани стандартным Ctrl-выбором AutoCAD.";

            if (!hasFaces || doc == null)
            {
                coatingBox.Text = string.Empty;
                RefreshCoatedFacesList(doc);
                return;
            }

            var values = new List<string>();
            using (Transaction tr = doc.Database.TransactionManager.StartTransaction())
            {
                foreach (FaceSelectionRef face in selectedFaces)
                {
                    var solid = tr.GetObject(face.SolidId, OpenMode.ForRead, false) as Solid3d;
                    if (solid == null || solid.IsErased)
                    {
                        continue;
                    }

                    values.Add(FaceCoatingStorage.Read(solid, tr, face.FaceKey));
                }

                tr.Commit();
            }

            coatingBox.Text = Merge(values);
            RefreshCoatedFacesList(doc);
        }

        private void ApplyCoatingToFaces(Document doc)
        {
            if (doc == null || selectedFaces.Count == 0 || IsMergedValue(coatingBox.Text))
            {
                return;
            }

            using (doc.LockDocument())
            using (Transaction tr = doc.Database.TransactionManager.StartTransaction())
            {
                int written = 0;
                foreach (FaceSelectionRef face in selectedFaces)
                {
                    var solid = tr.GetObject(face.SolidId, OpenMode.ForWrite, false) as Solid3d;
                    if (solid == null || solid.IsErased)
                    {
                        continue;
                    }

                    FaceCoatingStorage.Write(solid, tr, face.FaceKey, coatingBox.Text ?? string.Empty);
                    written++;
                }

                tr.Commit();
                coatingStatusText.Text = "Облицовка записана в граней: " + written.ToString(CultureInfo.InvariantCulture);
            }

            RefreshFaceBlock(doc);
        }

        private void RefreshCoatedFacesList(Document doc)
        {
            if (coatedFacesBox == null)
            {
                return;
            }

            if (doc == null || selectedSolidIds.Count == 0)
            {
                coatedFacesBox.Text = string.Empty;
                return;
            }

            var lines = new List<string>();
            using (Transaction tr = doc.Database.TransactionManager.StartTransaction())
            {
                foreach (ObjectId id in selectedSolidIds)
                {
                    var solid = tr.GetObject(id, OpenMode.ForRead, false) as Solid3d;
                    if (solid == null || solid.IsErased)
                    {
                        continue;
                    }

                    Dictionary<string, string> coatings = FaceCoatingStorage.ReadFaceCoatings(solid, tr);
                    foreach (KeyValuePair<string, string> coating in coatings.OrderBy(x => x.Value, StringComparer.CurrentCultureIgnoreCase))
                    {
                        if (string.IsNullOrWhiteSpace(coating.Value))
                        {
                            continue;
                        }

                        lines.Add(solid.Handle + "  "
                            + FormatFaceDimensions(coating.Key) + "  "
                            + coating.Value);
                    }
                }

                tr.Commit();
            }

            coatedFacesBox.Text = lines.Count == 0
                ? "Нет облицованных граней у выбранных тел."
                : string.Join(Environment.NewLine, lines);
        }

        private void SetSelectionStatus()
        {
            SetStatus("Выбрано 3D-тел: " + selectedSolidIds.Count.ToString(CultureInfo.InvariantCulture)
                + ", граней: " + selectedFaces.Count.ToString(CultureInfo.InvariantCulture));
        }

        private static string FormatFaceDimensions(string faceKey)
        {
            double first;
            double second;
            if (FaceKeyBuilder.TryGetBoundingDimensions(faceKey, out first, out second))
            {
                return Format(first) + " x " + Format(second);
            }

            return "габарит ?";
        }

        private static void ReadSelection(Database db, SelectionSet selectionSet, out List<ObjectId> solidIds, out List<FaceSelectionRef> faces)
        {
            solidIds = new List<ObjectId>();
            faces = new List<FaceSelectionRef>();

            RXClass solidClass = RXObject.GetClass(typeof(Solid3d));
            var selectedEdgesBySolid = new Dictionary<ObjectId, HashSet<string>>();
            int selectedIndex = 0;
            int faceCount = 0;
            int edgeCount = 0;
            int vertexCount = 0;
            int otherSubentityCount = 0;

            foreach (SelectedObject selected in selectionSet)
            {
                selectedIndex++;
                if (selected == null)
                {
                    continue;
                }

                ObjectId selectedSolidId = selected.ObjectId;
                if (!selectedSolidId.IsNull && selectedSolidId.ObjectClass == solidClass && !solidIds.Contains(selectedSolidId))
                {
                    solidIds.Add(selectedSolidId);
                }

                SelectedSubObject[] subentities = selected.GetSubentities();
                if (subentities == null)
                {
                    continue;
                }

                foreach (SelectedSubObject subentity in subentities)
                {
                    if (subentity == null)
                    {
                        continue;
                    }

                    FullSubentityPath path = subentity.FullSubentityPath;
                    if (path.IsNull)
                    {
                        continue;
                    }

                    if (path.SubentId.Type == SubentityType.Edge)
                    {
                        edgeCount++;
                        ObjectId edgeSolidId = GetSolidId(path, selectedSolidId);
                        if (!edgeSolidId.IsNull && edgeSolidId.ObjectClass == solidClass)
                        {
                            if (!solidIds.Contains(edgeSolidId))
                            {
                                solidIds.Add(edgeSolidId);
                            }

                            HashSet<string> edgeKeys;
                            if (!selectedEdgesBySolid.TryGetValue(edgeSolidId, out edgeKeys))
                            {
                                edgeKeys = new HashSet<string>(StringComparer.Ordinal);
                                selectedEdgesBySolid.Add(edgeSolidId, edgeKeys);
                            }

                            string edgeKey;
                            if (TryGetEdgeGeometryKey(path, out edgeKey))
                            {
                                edgeKeys.Add(edgeKey);
                            }
                            else
                            {
                                edgeKeys.Add(GetSubentityStableKey(path.SubentId));
                            }
                        }

                        continue;
                    }

                    if (path.SubentId.Type == SubentityType.Vertex)
                    {
                        vertexCount++;
                        continue;
                    }

                    if (path.SubentId.Type != SubentityType.Face)
                    {
                        otherSubentityCount++;
                        continue;
                    }

                    faceCount++;
                    ObjectId solidId = GetSolidId(path, selectedSolidId);
                    if (solidId.IsNull || solidId.ObjectClass != solidClass)
                    {
                        continue;
                    }

                    if (!solidIds.Contains(solidId))
                    {
                        solidIds.Add(solidId);
                    }

                    string key;
                    if (!FaceKeyBuilder.TryBuild(path, out key))
                    {
                        PaletteDebugLogger.Info(db, "ReadSelection: FaceKeyBuilder.TryBuild failed, using legacy key");
                        key = GetFaceKey(path.SubentId);
                    }

                    if (!faces.Any(x => x.SolidId == solidId && x.FaceKey == key))
                    {
                        faces.Add(new FaceSelectionRef(solidId, key));
                    }
                }
            }

            int inferredFaces = InferFacesFromSelectedEdges(db, solidClass, selectedEdgesBySolid, solidIds, faces);

            PaletteDebugLogger.Info(db, "ReadSelection summary: selected="
                + selectedIndex.ToString(CultureInfo.InvariantCulture)
                + " solids=" + solidIds.Count.ToString(CultureInfo.InvariantCulture)
                + " rawFaces=" + faceCount.ToString(CultureInfo.InvariantCulture)
                + " faces=" + faces.Count.ToString(CultureInfo.InvariantCulture)
                + " inferredFaces=" + inferredFaces.ToString(CultureInfo.InvariantCulture)
                + " edges=" + edgeCount.ToString(CultureInfo.InvariantCulture)
                + " vertices=" + vertexCount.ToString(CultureInfo.InvariantCulture)
                + " other=" + otherSubentityCount.ToString(CultureInfo.InvariantCulture));
        }

        private static int InferFacesFromSelectedEdges(Database db,
                                                       RXClass solidClass,
                                                       IDictionary<ObjectId, HashSet<string>> selectedEdgesBySolid,
                                                       IList<ObjectId> solidIds,
                                                       IList<FaceSelectionRef> faces)
        {
            if (db == null || selectedEdgesBySolid.Count == 0)
            {
                return 0;
            }

            int inferred = 0;

            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                foreach (KeyValuePair<ObjectId, HashSet<string>> pair in selectedEdgesBySolid)
                {
                    ObjectId solidId = pair.Key;
                    HashSet<string> selectedEdgeKeys = pair.Value;
                    if (solidId.IsNull || solidId.ObjectClass != solidClass || selectedEdgeKeys.Count == 0)
                    {
                        continue;
                    }

                    var solid = tr.GetObject(solidId, OpenMode.ForRead, false) as Solid3d;
                    if (solid == null || solid.IsErased)
                    {
                        continue;
                    }

                    try
                    {
                        using (var brep = new Brep(solid))
                        {
                            foreach (BrepFace face in brep.Faces)
                            {
                                HashSet<string> faceEdgeKeys = GetFaceExteriorEdgeKeys(face);
                                if (faceEdgeKeys.Count == 0)
                                {
                                    faceEdgeKeys = GetFaceAllEdgeKeys(face);
                                }

                                if (faceEdgeKeys.Count == 0 || faceEdgeKeys.Count > selectedEdgeKeys.Count)
                                {
                                    continue;
                                }

                                bool allEdgesSelected = faceEdgeKeys.All(selectedEdgeKeys.Contains);
                                if (!allEdgesSelected)
                                {
                                    continue;
                                }

                                string faceKey;
                                if (!FaceKeyBuilder.TryBuild(face, out faceKey))
                                {
                                    faceKey = GetFaceKey(face.SubentityPath.SubentId);
                                }

                                if (!faces.Any(x => x.SolidId == solidId && x.FaceKey == faceKey))
                                {
                                    faces.Add(new FaceSelectionRef(solidId, faceKey));
                                    inferred++;
                                }
                            }
                        }
                    }
                    catch (System.Exception ex)
                    {
                        PaletteDebugLogger.Error(db, "InferFacesFromSelectedEdges failed for solid " + FormatObjectId(solidId), ex);
                    }
                }

                tr.Commit();
            }

            return inferred;
        }

        private static HashSet<string> GetFaceExteriorEdgeKeys(BrepFace face)
        {
            var keys = new HashSet<string>(StringComparer.Ordinal);

            foreach (BoundaryLoop loop in face.Loops)
            {
                if (loop == null || loop.IsNull || loop.LoopType != LoopType.LoopExterior)
                {
                    continue;
                }

                AddLoopEdgeKeys(loop, keys);
            }

            return keys;
        }

        private static HashSet<string> GetFaceAllEdgeKeys(BrepFace face)
        {
            var keys = new HashSet<string>(StringComparer.Ordinal);

            foreach (BoundaryLoop loop in face.Loops)
            {
                if (loop == null || loop.IsNull)
                {
                    continue;
                }

                AddLoopEdgeKeys(loop, keys);
            }

            return keys;
        }

        private static void AddLoopEdgeKeys(BoundaryLoop loop, ISet<string> keys)
        {
            foreach (BrepEdge edge in loop.Edges)
            {
                if (edge == null || edge.IsNull)
                {
                    continue;
                }

                string edgeKey;
                if (!TryGetEdgeGeometryKey(edge, out edgeKey))
                {
                    continue;
                }

                keys.Add(edgeKey);
            }
        }

        private static bool TryGetEdgeGeometryKey(FullSubentityPath path, out string key)
        {
            key = null;

            if (path.IsNull || path.SubentId.Type != SubentityType.Edge)
            {
                return false;
            }

            try
            {
                using (var edge = new BrepEdge(path))
                {
                    return TryGetEdgeGeometryKey(edge, out key);
                }
            }
            catch
            {
                return false;
            }
        }

        private static bool TryGetEdgeGeometryKey(BrepEdge edge, out string key)
        {
            key = null;

            try
            {
                if (edge == null || edge.IsNull || edge.Vertex1 == null || edge.Vertex2 == null)
                {
                    return false;
                }

                string first = GetPointKey(edge.Vertex1.Point);
                string second = GetPointKey(edge.Vertex2.Point);

                if (string.CompareOrdinal(first, second) > 0)
                {
                    string temp = first;
                    first = second;
                    second = temp;
                }

                key = "EG:" + first + "|" + second;
                return true;
            }
            catch
            {
                return false;
            }
        }

        private static string GetPointKey(Point3d point)
        {
            return QuantizePointCoordinate(point.X)
                + ","
                + QuantizePointCoordinate(point.Y)
                + ","
                + QuantizePointCoordinate(point.Z);
        }

        private static string QuantizePointCoordinate(double value)
        {
            double rounded = Math.Round(value / 0.001) * 0.001;
            if (Math.Abs(rounded) < 0.0005)
            {
                rounded = 0.0;
            }

            return rounded.ToString("0.###", CultureInfo.InvariantCulture);
        }

        private static string FormatObjectId(ObjectId id)
        {
            if (id.IsNull)
            {
                return "<null>";
            }

            try
            {
                return id.Handle.ToString();
            }
            catch
            {
                return "<no handle>";
            }
        }

        private static string FormatObjectClass(ObjectId id)
        {
            if (id.IsNull)
            {
                return "<null>";
            }

            try
            {
                RXClass objectClass = id.ObjectClass;
                return objectClass != null ? objectClass.Name : "<null class>";
            }
            catch (System.Exception ex)
            {
                return "<class error: " + ex.Message + ">";
            }
        }

        private static ObjectId GetSolidId(FullSubentityPath path, ObjectId fallback)
        {
            ObjectId[] ids = path.GetObjectIds();
            return ids != null && ids.Length > 0 ? ids[ids.Length - 1] : fallback;
        }

        private static string GetFaceKey(SubentityId subentityId)
        {
            return ((int)subentityId.Type).ToString(CultureInfo.InvariantCulture)
                + ":"
                + subentityId.Index.ToString(CultureInfo.InvariantCulture)
                + ":"
                + subentityId.IndexPtr.ToInt64().ToString(CultureInfo.InvariantCulture);
        }

        private static string GetSubentityStableKey(SubentityId subentityId)
        {
            return ((int)subentityId.Type).ToString(CultureInfo.InvariantCulture)
                + ":"
                + subentityId.IndexPtr.ToInt64().ToString(CultureInfo.InvariantCulture);
        }

        private static TextBox AddTextRow(Grid grid, ref int row, string label)
        {
            var box = new TextBox { Margin = FieldMargin() };
            AddRow(grid, ref row, label, box);
            return box;
        }

        private static void AddRow(Grid grid, ref int row, string label, UIElement field)
        {
            grid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });

            var labelBlock = new TextBlock
            {
                Text = label,
                Foreground = Brushes.Gainsboro,
                Margin = new Thickness(0, 7, 8, 4),
                VerticalAlignment = VerticalAlignment.Center
            };

            Grid.SetRow(labelBlock, row);
            Grid.SetColumn(labelBlock, 0);
            grid.Children.Add(labelBlock);

            Grid.SetRow(field, row);
            Grid.SetColumn(field, 1);
            grid.Children.Add(field);

            row++;
        }

        private static Thickness FieldMargin()
        {
            return new Thickness(0, 3, 0, 4);
        }

        private void SetFieldsEnabled(bool enabled)
        {
            assemblyNumberBox.IsEnabled = enabled;
            partNumberBox.IsEnabled = enabled;
            partNameBox.IsEnabled = enabled;
            partTypeBox.IsEnabled = enabled;
            lengthBox.IsEnabled = enabled;
            widthBox.IsEnabled = enabled;
            rotateBox.IsEnabled = enabled;
            materialBox.IsEnabled = enabled;
            noteBox.IsEnabled = enabled;
            statusText.IsEnabled = true;
            coatingPanel.IsEnabled = true;
            coatingBox.IsEnabled = selectedFaces.Count > 0;
            coatedFacesBox.IsEnabled = true;
        }

        private void ClearFields()
        {
            assemblyNumberBox.Text = string.Empty;
            partNumberBox.Text = string.Empty;
            partNameBox.Text = string.Empty;
            partTypeBox.SelectedItem = null;
            lengthBox.Text = string.Empty;
            widthBox.Text = string.Empty;
            rotateBox.IsChecked = false;
            materialBox.Text = string.Empty;
            noteBox.Text = string.Empty;
        }

        private void ClearSelectionState()
        {
            selectedSolidIds = new List<ObjectId>();
            selectedFaces = new List<FaceSelectionRef>();
            ClearFields();
            coatingBox.Text = string.Empty;
            coatedFacesBox.Text = string.Empty;
        }

        private void SetStatus(string text)
        {
            statusText.Text = text;
        }

        private static string Merge(IEnumerable<string> values)
        {
            var distinct = values.Select(x => x ?? string.Empty).Distinct(StringComparer.Ordinal).Take(2).ToList();
            return distinct.Count <= 1 ? distinct.FirstOrDefault() ?? string.Empty : "*РАЗНЫЕ*";
        }

        private void SetComboMerged(ComboBox comboBox, IEnumerable<string> values)
        {
            string value = Merge(values);
            if (value == "*РАЗНЫЕ*" || string.IsNullOrWhiteSpace(value))
            {
                comboBox.SelectedItem = null;
                comboBox.Text = value;
                return;
            }

            comboBox.SelectedItem = value;
        }

        private void SetCheckMerged(IEnumerable<bool> values)
        {
            var distinct = values.Distinct().Take(2).ToList();
            rotateBox.IsThreeState = distinct.Count > 1;
            rotateBox.IsChecked = distinct.Count > 1 ? (bool?)null : distinct.FirstOrDefault();
        }

        private static bool IsMergedValue(string value)
        {
            return value == "*РАЗНЫЕ*";
        }

        private static string FieldValue(string value, string existing)
        {
            return IsMergedValue(value) ? existing ?? string.Empty : value ?? string.Empty;
        }

        private static bool TryParseLength(string value, out double result)
        {
            value = (value ?? string.Empty).Replace(',', '.');
            return double.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out result);
        }

        private static string Format(double value)
        {
            return value.ToString("0.0", CultureInfo.InvariantCulture);
        }

        private static string FirstNonBlank(string value, string fallback)
        {
            return string.IsNullOrWhiteSpace(value) ? fallback : value;
        }

        private sealed class FaceSelectionRef
        {
            public FaceSelectionRef(ObjectId solidId, string faceKey)
            {
                SolidId = solidId;
                FaceKey = faceKey;
            }

            public ObjectId SolidId { get; private set; }
            public string FaceKey { get; private set; }
        }
    }
}
