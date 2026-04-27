using System;
using System.ComponentModel;
using System.Globalization;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Data;
using System.Windows.Media;
using AutoCAD_BoardSorter.Models;

namespace AutoCAD_BoardSorter.Ui
{
    internal sealed class MaterialDatabaseEditorWindow : Window
    {
        private readonly MaterialDatabase database;
        private readonly ListBox boardList = new ListBox();
        private readonly ListBox coatingList = new ListBox();
        private readonly ListBox formatList = new ListBox();
        private readonly ListBox boardCategoryList = new ListBox();
        private readonly ListBox coatingCategoryList = new ListBox();
        private readonly TreeView boardCategoryTree = new TreeView();
        private readonly TreeView coatingCategoryTree = new TreeView();
        private readonly ICollectionView boardMaterialsView;
        private readonly ICollectionView coatingMaterialsView;
        private readonly TextBox boardCodeBox = CreateTextBox();
        private readonly TextBox boardNameBox = CreateTextBox();
        private readonly ComboBox boardCalcBox = CreateCombo("Площадной", "Погонный", "Объемный", "Штучный");
        private readonly TextBox boardCategoryBox = CreateTextBox();
        private readonly ComboBox visibleEdgeBox = CreateMaterialComboBox();
        private readonly ComboBox hiddenEdgeBox = CreateMaterialComboBox();
        private readonly ComboBox frontFaceBox = CreateMaterialComboBox();
        private readonly ComboBox backFaceBox = CreateMaterialComboBox();
        private readonly TextBox formatCodeBox = CreateTextBox();
        private readonly ComboBox formatTypeBox = CreateCombo("Площадной", "Погонный", "Объемный");
        private readonly TextBox formatLengthBox = CreateTextBox();
        private readonly TextBox formatWidthBox = CreateTextBox();
        private readonly TextBox formatThicknessBox = CreateTextBox();
        private readonly TextBox coatingCodeBox = CreateTextBox();
        private readonly TextBox coatingNameBox = CreateTextBox();
        private readonly ComboBox coatingCalcBox = CreateCombo("Погонный", "Площадной");
        private readonly TextBox coatingThicknessBox = CreateTextBox();
        private readonly TextBox coatingCategoryBox = CreateTextBox();
        private readonly TextBox boardCategoryCodeBox = CreateTextBox();
        private readonly TextBox boardCategoryNameBox = CreateTextBox();
        private readonly TextBox boardCategoryParentBox = CreateTextBox();
        private readonly TextBox coatingCategoryCodeBox = CreateTextBox();
        private readonly TextBox coatingCategoryNameBox = CreateTextBox();
        private readonly TextBox coatingCategoryParentBox = CreateTextBox();
        private BoardMaterialData currentBoard;
        private MaterialFormatData currentFormat;
        private CoatingMaterialData currentCoating;
        private MaterialCategoryData currentBoardCategory;
        private MaterialCategoryData currentCoatingCategory;
        private bool updating;
        private string selectedBoardCategoryCode;
        private string selectedCoatingCategoryCode;
        private Grid boardLeftGrid;
        private Grid coatingLeftGrid;
        private ColumnDefinition boardSidebarColumn;
        private ColumnDefinition coatingSidebarColumn;
        private RowDefinition boardTreeRow;
        private RowDefinition boardListRow;
        private RowDefinition coatingTreeRow;
        private RowDefinition coatingListRow;

        public MaterialDatabaseEditorWindow()
        {
            database = MaterialDatabaseStore.Load();
            boardMaterialsView = CollectionViewSource.GetDefaultView(database.BoardMaterials);
            coatingMaterialsView = CollectionViewSource.GetDefaultView(database.CoatingMaterials);
            Title = "База материалов";
            Width = 980;
            Height = 720;
            WindowStartupLocation = WindowStartupLocation.CenterScreen;
            Background = new SolidColorBrush(Color.FromRgb(15, 23, 42));
            Foreground = Brushes.White;
            FontFamily = new FontFamily("Segoe UI");
            Resources[typeof(TextBox)] = CreateTextBoxStyle();
            Resources[typeof(ComboBox)] = CreateComboStyle();
            Resources[typeof(ComboBoxItem)] = CreateComboBoxItemStyle();
            Resources[typeof(ListBox)] = CreateListBoxStyle();
            Resources[typeof(ListBoxItem)] = CreateListBoxItemStyle();
            Resources[typeof(TreeView)] = CreateTreeViewStyle();
            Resources[typeof(TreeViewItem)] = CreateTreeViewItemStyle();
            Resources[typeof(TabControl)] = CreateTabControlStyle();
            Resources[typeof(TabItem)] = CreateTabItemStyle();
            Resources[SystemColors.WindowBrushKey] = new SolidColorBrush(Color.FromRgb(15, 23, 42));
            Resources[SystemColors.ControlBrushKey] = new SolidColorBrush(Color.FromRgb(17, 24, 39));
            Resources[SystemColors.ControlLightBrushKey] = new SolidColorBrush(Color.FromRgb(17, 24, 39));
            Resources[SystemColors.ControlDarkBrushKey] = new SolidColorBrush(Color.FromRgb(51, 65, 85));
            Resources[SystemColors.HighlightBrushKey] = new SolidColorBrush(Color.FromRgb(37, 99, 235));
            Resources[SystemColors.HighlightTextBrushKey] = Brushes.White;
            Resources[SystemColors.InactiveSelectionHighlightBrushKey] = new SolidColorBrush(Color.FromRgb(37, 99, 235));
            Resources[SystemColors.InactiveSelectionHighlightTextBrushKey] = Brushes.White;
            boardMaterialsView.Filter = FilterBoardMaterial;
            coatingMaterialsView.Filter = FilterCoatingMaterial;
            Closing += delegate { CaptureLayoutToDatabase(); MaterialDatabaseStore.Save(database); };
            Content = BuildUi();
            RefreshAll();
        }

        private UIElement BuildUi()
        {
            var shell = new Border
            {
                Background = new SolidColorBrush(Color.FromRgb(15, 23, 42)),
                Padding = new Thickness(14)
            };
            var root = new DockPanel();
            var footer = new StackPanel { Orientation = Orientation.Horizontal, HorizontalAlignment = HorizontalAlignment.Right, Margin = new Thickness(0, 12, 0, 0) };
            var save = new Button { Content = "Сохранить", Padding = new Thickness(18, 8, 18, 8), Margin = new Thickness(0, 0, 8, 0) };
            var close = new Button { Content = "Закрыть", Padding = new Thickness(18, 8, 18, 8) };
            ApplyButtonStyle(save);
            ApplyButtonStyle(close);
            save.Click += delegate { SaveCurrent(); CaptureLayoutToDatabase(); MaterialDatabaseStore.Save(database); Close(); };
            close.Click += delegate { Close(); };
            footer.Children.Add(save);
            footer.Children.Add(close);
            DockPanel.SetDock(footer, Dock.Bottom);
            root.Children.Add(footer);

            var tabs = new TabControl();
            tabs.Items.Add(new TabItem { Header = "Листовые материалы", Content = BuildBoardTab() });
            tabs.Items.Add(new TabItem { Header = "Облицовочные материалы", Content = BuildCoatingTab() });
            tabs.Items.Add(new TabItem { Header = "Категории", Content = BuildCategoriesTab() });
            root.Children.Add(tabs);
            shell.Child = root;
            return shell;
        }

        private UIElement BuildBoardTab()
        {
            var grid = new Grid();
            boardSidebarColumn = new ColumnDefinition { Width = new GridLength(Math.Max(180.0, database.UiLayout.BoardSidebarWidth)) };
            grid.ColumnDefinitions.Add(boardSidebarColumn);
            grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(6) });
            grid.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
            grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(6) });
            grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });

            boardLeftGrid = new Grid();
            boardTreeRow = new RowDefinition { Height = new GridLength(Math.Max(120.0, database.UiLayout.BoardTreeHeight), GridUnitType.Pixel) };
            boardListRow = new RowDefinition { Height = new GridLength(1, GridUnitType.Star) };
            boardLeftGrid.RowDefinitions.Add(boardTreeRow);
            boardLeftGrid.RowDefinitions.Add(new RowDefinition { Height = new GridLength(6) });
            boardLeftGrid.RowDefinitions.Add(boardListRow);

            var treeHost = BuildCategoryTreePanel(boardCategoryTree, "Категории", AddBoardCategory, DeleteBoardCategory, true, OnBoardCategoryTreeSelectionChanged);
            Grid.SetRow(treeHost, 0);
            boardLeftGrid.Children.Add(treeHost);
            var treeSplitter = new GridSplitter
            {
                Height = 6,
                ResizeDirection = GridResizeDirection.Rows,
                ResizeBehavior = GridResizeBehavior.PreviousAndNext,
                HorizontalAlignment = HorizontalAlignment.Stretch,
                VerticalAlignment = VerticalAlignment.Stretch,
                Background = new SolidColorBrush(Color.FromRgb(51, 65, 85))
            };
            Grid.SetRow(treeSplitter, 1);
            boardLeftGrid.Children.Add(treeSplitter);
            var listHost = WrapList(boardList, "Материалы", AddBoard, DeleteBoard);
            Grid.SetRow(listHost, 2);
            boardLeftGrid.Children.Add(listHost);

            Grid.SetColumn(boardLeftGrid, 0);
            grid.Children.Add(boardLeftGrid);

            var leftSplitter = new GridSplitter
            {
                Width = 6,
                ResizeDirection = GridResizeDirection.Columns,
                ResizeBehavior = GridResizeBehavior.PreviousAndNext,
                HorizontalAlignment = HorizontalAlignment.Stretch,
                VerticalAlignment = VerticalAlignment.Stretch,
                Background = new SolidColorBrush(Color.FromRgb(51, 65, 85))
            };
            Grid.SetColumn(leftSplitter, 2);
            grid.Children.Add(leftSplitter);

            boardList.SelectionChanged += delegate { OnBoardSelectionChanged(); };

            var panel = new ScrollViewer { VerticalScrollBarVisibility = ScrollBarVisibility.Auto };
            var stack = Stack();
            panel.Content = stack;
            stack.Children.Add(Label("ID Код"));
            stack.Children.Add(boardCodeBox);
            stack.Children.Add(Label("Наименование"));
            stack.Children.Add(boardNameBox);
            stack.Children.Add(Label("Тип расчета"));
            stack.Children.Add(boardCalcBox);
            stack.Children.Add(Label("Категория"));
            stack.Children.Add(boardCategoryBox);
            stack.Children.Add(Label("Стандартная облицовка торцев лицевая"));
            stack.Children.Add(visibleEdgeBox);
            stack.Children.Add(Label("Стандартная облицовка торцев не видимая"));
            stack.Children.Add(hiddenEdgeBox);
            stack.Children.Add(Label("Облицовка пласти лицевая"));
            stack.Children.Add(frontFaceBox);
            stack.Children.Add(Label("Облицовка пласти тыльная"));
            stack.Children.Add(backFaceBox);
            stack.Children.Add(Label("Форматы"));
            stack.Children.Add(WrapList(formatList, "Форматы", AddFormat, DeleteFormat));
            formatList.SelectionChanged += delegate { OnFormatSelectionChanged(); };
            stack.Children.Add(Label("Формат: ID Код"));
            stack.Children.Add(formatCodeBox);
            stack.Children.Add(Label("Формат: тип"));
            stack.Children.Add(formatTypeBox);
            stack.Children.Add(Label("Длина / Ширина / Толщина"));
            var dims = new Grid();
            dims.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
            dims.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
            dims.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
            Grid.SetColumn(formatLengthBox, 0);
            Grid.SetColumn(formatWidthBox, 1);
            Grid.SetColumn(formatThicknessBox, 2);
            dims.Children.Add(formatLengthBox);
            dims.Children.Add(formatWidthBox);
            dims.Children.Add(formatThicknessBox);
            stack.Children.Add(dims);
            Grid.SetColumn(panel, 4);
            grid.Children.Add(panel);
            return grid;
        }

        private UIElement BuildCoatingTab()
        {
            var grid = new Grid();
            coatingSidebarColumn = new ColumnDefinition { Width = new GridLength(Math.Max(180.0, database.UiLayout.CoatingSidebarWidth)) };
            grid.ColumnDefinitions.Add(coatingSidebarColumn);
            grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(6) });
            grid.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
            grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(6) });
            grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });

            coatingLeftGrid = new Grid();
            coatingTreeRow = new RowDefinition { Height = new GridLength(Math.Max(120.0, database.UiLayout.CoatingTreeHeight), GridUnitType.Pixel) };
            coatingListRow = new RowDefinition { Height = new GridLength(1, GridUnitType.Star) };
            coatingLeftGrid.RowDefinitions.Add(coatingTreeRow);
            coatingLeftGrid.RowDefinitions.Add(new RowDefinition { Height = new GridLength(6) });
            coatingLeftGrid.RowDefinitions.Add(coatingListRow);

            var treeHost = BuildCategoryTreePanel(coatingCategoryTree, "Категории", AddCoatingCategory, DeleteCoatingCategory, false, OnCoatingCategoryTreeSelectionChanged);
            Grid.SetRow(treeHost, 0);
            coatingLeftGrid.Children.Add(treeHost);
            var treeSplitter = new GridSplitter
            {
                Height = 6,
                ResizeDirection = GridResizeDirection.Rows,
                ResizeBehavior = GridResizeBehavior.PreviousAndNext,
                HorizontalAlignment = HorizontalAlignment.Stretch,
                VerticalAlignment = VerticalAlignment.Stretch,
                Background = new SolidColorBrush(Color.FromRgb(51, 65, 85))
            };
            Grid.SetRow(treeSplitter, 1);
            coatingLeftGrid.Children.Add(treeSplitter);
            var listHost = WrapList(coatingList, "Материалы", AddCoating, DeleteCoating);
            Grid.SetRow(listHost, 2);
            coatingLeftGrid.Children.Add(listHost);

            Grid.SetColumn(coatingLeftGrid, 0);
            grid.Children.Add(coatingLeftGrid);

            var leftSplitter = new GridSplitter
            {
                Width = 6,
                ResizeDirection = GridResizeDirection.Columns,
                ResizeBehavior = GridResizeBehavior.PreviousAndNext,
                HorizontalAlignment = HorizontalAlignment.Stretch,
                VerticalAlignment = VerticalAlignment.Stretch,
                Background = new SolidColorBrush(Color.FromRgb(51, 65, 85))
            };
            Grid.SetColumn(leftSplitter, 2);
            grid.Children.Add(leftSplitter);

            coatingList.SelectionChanged += delegate { OnCoatingSelectionChanged(); };

            var stack = Stack();
            stack.Children.Add(Label("ID Код"));
            stack.Children.Add(coatingCodeBox);
            stack.Children.Add(Label("Наименование"));
            stack.Children.Add(coatingNameBox);
            stack.Children.Add(Label("Тип расчета"));
            stack.Children.Add(coatingCalcBox);
            stack.Children.Add(Label("Толщина"));
            stack.Children.Add(coatingThicknessBox);
            stack.Children.Add(Label("Категория"));
            stack.Children.Add(coatingCategoryBox);
            Grid.SetColumn(stack, 4);
            grid.Children.Add(stack);
            return grid;
        }

        private UIElement BuildCategoryTreePanel(TreeView tree, string title, Action add, Action delete, bool boardTree, RoutedPropertyChangedEventHandler<object> selectionChanged)
        {
            var dock = new DockPanel { Margin = new Thickness(0, 0, 12, 0) };
            var header = new StackPanel { Orientation = Orientation.Horizontal, Margin = new Thickness(0, 0, 0, 8) };
            header.Children.Add(new TextBlock { Text = title, Foreground = Brushes.White, FontWeight = FontWeights.SemiBold, VerticalAlignment = VerticalAlignment.Center });
            var addButton = new Button { Content = "+", Width = 28, Margin = new Thickness(8, 0, 4, 0) };
            var delButton = new Button { Content = "-", Width = 28 };
            ApplyButtonStyle(addButton);
            ApplyButtonStyle(delButton);
            addButton.Click += delegate { add(); };
            delButton.Click += delegate { delete(); };
            header.Children.Add(addButton);
            header.Children.Add(delButton);
            DockPanel.SetDock(header, Dock.Top);
            dock.Children.Add(header);

            tree.SelectedItemChanged += selectionChanged;
            dock.Children.Add(tree);
            return dock;
        }

        private void OnBoardCategoryTreeSelectionChanged(object sender, RoutedPropertyChangedEventArgs<object> e)
        {
            if (updating)
            {
                return;
            }

            var item = e.NewValue as TreeViewItem;
            var category = item != null ? item.Tag as MaterialCategoryData : null;
            selectedBoardCategoryCode = category != null ? category.Code : null;
            boardMaterialsView.Refresh();
        }

        private void OnCoatingCategoryTreeSelectionChanged(object sender, RoutedPropertyChangedEventArgs<object> e)
        {
            if (updating)
            {
                return;
            }

            var item = e.NewValue as TreeViewItem;
            var category = item != null ? item.Tag as MaterialCategoryData : null;
            selectedCoatingCategoryCode = category != null ? category.Code : null;
            coatingMaterialsView.Refresh();
        }

        private UIElement BuildCategoriesTab()
        {
            var grid = new Grid();
            grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
            grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });

            var board = BuildCategoryEditor(
                "Категории листовых материалов",
                boardCategoryList,
                boardCategoryCodeBox,
                boardCategoryNameBox,
                boardCategoryParentBox,
                AddBoardCategory,
                DeleteBoardCategory);
            boardCategoryList.SelectionChanged += delegate { OnBoardCategorySelectionChanged(); };
            Grid.SetColumn(board, 0);
            grid.Children.Add(board);

            var coating = BuildCategoryEditor(
                "Категории облицовочных материалов",
                coatingCategoryList,
                coatingCategoryCodeBox,
                coatingCategoryNameBox,
                coatingCategoryParentBox,
                AddCoatingCategory,
                DeleteCoatingCategory);
            coatingCategoryList.SelectionChanged += delegate { OnCoatingCategorySelectionChanged(); };
            Grid.SetColumn(coating, 1);
            grid.Children.Add(coating);

            return grid;
        }

        private UIElement BuildCategoryEditor(
            string title,
            ListBox list,
            TextBox codeBox,
            TextBox nameBox,
            TextBox parentBox,
            Action add,
            Action delete)
        {
            var dock = new DockPanel { Margin = new Thickness(0, 0, 12, 0) };
            list.Width = 260;
            UIElement wrappedList = WrapList(list, title, add, delete);
            DockPanel.SetDock(wrappedList, Dock.Left);
            dock.Children.Add(wrappedList);

            var stack = Stack();
            stack.Children.Add(Label("ID Код"));
            stack.Children.Add(codeBox);
            stack.Children.Add(Label("Наименование"));
            stack.Children.Add(nameBox);
            stack.Children.Add(Label("Родительская категория"));
            stack.Children.Add(parentBox);
            stack.Children.Add(new TextBox
            {
                Text = "Дерево строится по полю \"Родительская категория\".\n"
                    + "Чтобы сделать подкатегорию, укажи код родителя.\n\n"
                    + "Файл базы:\n" + MaterialDatabaseStore.FilePath,
                IsReadOnly = true,
                TextWrapping = TextWrapping.Wrap,
                Background = new SolidColorBrush(Color.FromRgb(17, 24, 39)),
                Foreground = Brushes.White,
                BorderBrush = new SolidColorBrush(Color.FromRgb(51, 65, 85)),
                Padding = new Thickness(12)
            });
            dock.Children.Add(stack);
            return dock;
        }

        private void AddBoard()
        {
            SaveCurrent();
            var material = new BoardMaterialData { Code = UniqueCode("MAT", database.BoardMaterials.Select(x => x.Code)), Name = "Новый материал", CalculationType = "Площадной", CategoryCode = selectedBoardCategoryCode ?? string.Empty };
            database.BoardMaterials.Add(material);
            RefreshAll();
            boardList.SelectedItem = material;
        }

        private void DeleteBoard()
        {
            var material = boardList.SelectedItem as BoardMaterialData;
            if (material == null) return;
            if (!ConfirmDelete("материал", material.Name ?? material.Code))
            {
                return;
            }
            database.BoardMaterials.Remove(material);
            RefreshAll();
        }

        private void AddFormat()
        {
            SaveCurrent();
            var material = boardList.SelectedItem as BoardMaterialData;
            if (material == null) return;
            var format = new MaterialFormatData { Code = UniqueCode("FMT", material.Formats.Select(x => x.Code)), FormatType = "Площадной", Length = 2800, Width = 2070, Thickness = 16 };
            material.Formats.Add(format);
            RefreshFormats(material);
            formatList.SelectedItem = format;
        }

        private void DeleteFormat()
        {
            var material = boardList.SelectedItem as BoardMaterialData;
            var format = formatList.SelectedItem as MaterialFormatData;
            if (material == null || format == null) return;
            material.Formats.Remove(format);
            RefreshFormats(material);
        }

        private void AddCoating()
        {
            SaveCurrent();
            var material = new CoatingMaterialData { Code = UniqueCode("COAT", database.CoatingMaterials.Select(x => x.Code)), Name = "Новая облицовка", CalculationType = "Погонный", CategoryCode = selectedCoatingCategoryCode ?? string.Empty };
            database.CoatingMaterials.Add(material);
            RefreshAll();
            coatingList.SelectedItem = material;
        }

        private void DeleteCoating()
        {
            var material = coatingList.SelectedItem as CoatingMaterialData;
            if (material == null) return;
            if (!ConfirmDelete("облицовочный материал", material.Name ?? material.Code))
            {
                return;
            }
            database.CoatingMaterials.Remove(material);
            RefreshAll();
        }

        private void AddBoardCategory()
        {
            SaveCurrent();
            var category = new MaterialCategoryData { Code = UniqueCode("CAT", database.BoardCategories.Select(x => x.Code)), Name = "Новая категория" };
            database.BoardCategories.Add(category);
            RefreshCategories();
            boardCategoryList.SelectedItem = category;
        }

        private void DeleteBoardCategory()
        {
            var category = boardCategoryList.SelectedItem as MaterialCategoryData;
            if (category == null) return;
            if (!ConfirmDelete("категорию", category.Name ?? category.Code))
            {
                return;
            }
            database.BoardCategories.Remove(category);
            selectedBoardCategoryCode = null;
            RefreshCategories();
        }

        private void AddCoatingCategory()
        {
            SaveCurrent();
            var category = new MaterialCategoryData { Code = UniqueCode("CAT", database.CoatingCategories.Select(x => x.Code)), Name = "Новая категория" };
            database.CoatingCategories.Add(category);
            RefreshCategories();
            coatingCategoryList.SelectedItem = category;
        }

        private void DeleteCoatingCategory()
        {
            var category = coatingCategoryList.SelectedItem as MaterialCategoryData;
            if (category == null) return;
            if (!ConfirmDelete("категорию", category.Name ?? category.Code))
            {
                return;
            }
            database.CoatingCategories.Remove(category);
            selectedCoatingCategoryCode = null;
            RefreshCategories();
        }

        private void OnBoardSelectionChanged()
        {
            if (updating) return;
            SaveCurrent();
            currentBoard = boardList.SelectedItem as BoardMaterialData;
            currentFormat = null;
            LoadSelectedBoard();
        }

        private void OnFormatSelectionChanged()
        {
            if (updating) return;
            SaveCurrent();
            currentFormat = formatList.SelectedItem as MaterialFormatData;
            LoadSelectedFormat();
        }

        private void OnCoatingSelectionChanged()
        {
            if (updating) return;
            SaveCurrent();
            currentCoating = coatingList.SelectedItem as CoatingMaterialData;
            LoadSelectedCoating();
        }

        private void OnBoardCategorySelectionChanged()
        {
            if (updating) return;
            SaveCurrent();
            currentBoardCategory = boardCategoryList.SelectedItem as MaterialCategoryData;
            LoadSelectedBoardCategory();
        }

        private void OnCoatingCategorySelectionChanged()
        {
            if (updating) return;
            SaveCurrent();
            currentCoatingCategory = coatingCategoryList.SelectedItem as MaterialCategoryData;
            LoadSelectedCoatingCategory();
        }

        private void LoadSelectedBoard()
        {
            if (updating) return;
            updating = true;
            try
            {
                var material = boardList.SelectedItem as BoardMaterialData;
                currentBoard = material;
                boardCodeBox.Text = material?.Code ?? string.Empty;
                boardNameBox.Text = material?.Name ?? string.Empty;
                boardCalcBox.SelectedItem = material != null ? material.CalculationType ?? "Площадной" : "Площадной";
                boardCategoryBox.Text = material?.CategoryCode ?? string.Empty;
                SelectCoatingCombo(visibleEdgeBox, material?.DefaultVisibleEdgeCoatingCode);
                SelectCoatingCombo(hiddenEdgeBox, material?.DefaultHiddenEdgeCoatingCode);
                SelectCoatingCombo(frontFaceBox, material?.FrontFaceCoatingCode);
                SelectCoatingCombo(backFaceBox, material?.BackFaceCoatingCode);
                RefreshFormats(material);
            }
            finally { updating = false; }
        }

        private void LoadSelectedFormat()
        {
            if (updating) return;
            updating = true;
            try
            {
                var format = formatList.SelectedItem as MaterialFormatData;
                currentFormat = format;
                formatCodeBox.Text = format?.Code ?? string.Empty;
                formatTypeBox.SelectedItem = format != null ? format.FormatType ?? "Площадной" : "Площадной";
                formatLengthBox.Text = format != null ? format.Length.ToString("0.###", CultureInfo.InvariantCulture) : string.Empty;
                formatWidthBox.Text = format != null ? format.Width.ToString("0.###", CultureInfo.InvariantCulture) : string.Empty;
                formatThicknessBox.Text = format != null ? format.Thickness.ToString("0.###", CultureInfo.InvariantCulture) : string.Empty;
            }
            finally { updating = false; }
        }

        private void LoadSelectedCoating()
        {
            if (updating) return;
            updating = true;
            try
            {
                var material = coatingList.SelectedItem as CoatingMaterialData;
                currentCoating = material;
                coatingCodeBox.Text = material?.Code ?? string.Empty;
                coatingNameBox.Text = material?.Name ?? string.Empty;
                coatingCalcBox.SelectedItem = material != null ? material.CalculationType ?? "Погонный" : "Погонный";
                coatingThicknessBox.Text = material != null ? material.Thickness.ToString("0.###", CultureInfo.InvariantCulture) : string.Empty;
                coatingCategoryBox.Text = material?.CategoryCode ?? string.Empty;
            }
            finally { updating = false; }
        }

        private void LoadSelectedBoardCategory()
        {
            if (updating) return;
            updating = true;
            try
            {
                var category = boardCategoryList.SelectedItem as MaterialCategoryData;
                currentBoardCategory = category;
                boardCategoryCodeBox.Text = category?.Code ?? string.Empty;
                boardCategoryNameBox.Text = category?.Name ?? string.Empty;
                boardCategoryParentBox.Text = category?.ParentCode ?? string.Empty;
            }
            finally { updating = false; }
        }

        private void LoadSelectedCoatingCategory()
        {
            if (updating) return;
            updating = true;
            try
            {
                var category = coatingCategoryList.SelectedItem as MaterialCategoryData;
                currentCoatingCategory = category;
                coatingCategoryCodeBox.Text = category?.Code ?? string.Empty;
                coatingCategoryNameBox.Text = category?.Name ?? string.Empty;
                coatingCategoryParentBox.Text = category?.ParentCode ?? string.Empty;
            }
            finally { updating = false; }
        }

        private void SaveCurrent()
        {
            var board = currentBoard;
            if (board != null)
            {
                board.Code = boardCodeBox.Text.Trim();
                board.Name = boardNameBox.Text.Trim();
                board.CalculationType = SelectedComboText(boardCalcBox, "Площадной");
                board.CategoryCode = boardCategoryBox.Text.Trim();
                board.DefaultVisibleEdgeCoatingCode = SelectedCoatingCode(visibleEdgeBox);
                board.DefaultHiddenEdgeCoatingCode = SelectedCoatingCode(hiddenEdgeBox);
                board.FrontFaceCoatingCode = SelectedCoatingCode(frontFaceBox);
                board.BackFaceCoatingCode = SelectedCoatingCode(backFaceBox);
            }

            var format = currentFormat;
            if (format != null)
            {
                format.Code = formatCodeBox.Text.Trim();
                format.FormatType = SelectedComboText(formatTypeBox, "Площадной");
                if (TryDouble(formatLengthBox.Text, out double length)) format.Length = length;
                if (TryDouble(formatWidthBox.Text, out double width)) format.Width = width;
                if (TryDouble(formatThicknessBox.Text, out double thickness)) format.Thickness = thickness;
            }

            var coating = currentCoating;
            if (coating != null)
            {
                coating.Code = coatingCodeBox.Text.Trim();
                coating.Name = coatingNameBox.Text.Trim();
                coating.CalculationType = SelectedComboText(coatingCalcBox, "Погонный");
                if (TryDouble(coatingThicknessBox.Text, out double thickness)) coating.Thickness = thickness;
                coating.CategoryCode = coatingCategoryBox.Text.Trim();
            }

            var boardCategory = currentBoardCategory;
            if (boardCategory != null)
            {
                boardCategory.Code = boardCategoryCodeBox.Text.Trim();
                boardCategory.Name = boardCategoryNameBox.Text.Trim();
                boardCategory.ParentCode = boardCategoryParentBox.Text.Trim();
            }

            var coatingCategory = currentCoatingCategory;
            if (coatingCategory != null)
            {
                coatingCategory.Code = coatingCategoryCodeBox.Text.Trim();
                coatingCategory.Name = coatingCategoryNameBox.Text.Trim();
                coatingCategory.ParentCode = coatingCategoryParentBox.Text.Trim();
            }
        }

        private void RefreshAll()
        {
            updating = true;
            try
            {
                RebuildCategoryTrees();
                RefreshMaterialViews();
                boardCategoryList.ItemsSource = null;
                boardCategoryList.ItemsSource = database.BoardCategories;
                coatingCategoryList.ItemsSource = null;
                coatingCategoryList.ItemsSource = database.CoatingCategories;
                var coatingItems = database.CoatingMaterials.ToList();
                visibleEdgeBox.ItemsSource = coatingItems;
                hiddenEdgeBox.ItemsSource = coatingItems;
                frontFaceBox.ItemsSource = coatingItems;
                backFaceBox.ItemsSource = coatingItems;
            }
            finally { updating = false; }
            LoadSelectedBoard();
            LoadSelectedCoating();
            LoadSelectedBoardCategory();
            LoadSelectedCoatingCategory();
            ApplyLayoutFromDatabase();
        }

        private void RebuildCategoryTrees()
        {
            boardCategoryTree.Items.Clear();
            coatingCategoryTree.Items.Clear();
            AddCategoryTreeNodes(boardCategoryTree, database.BoardCategories);
            AddCategoryTreeNodes(coatingCategoryTree, database.CoatingCategories);
        }

        private static void AddCategoryTreeNodes(TreeView tree, System.Collections.Generic.IEnumerable<MaterialCategoryData> categories)
        {
            if (tree == null)
            {
                return;
            }

            var list = categories != null ? categories.Where(x => x != null).ToList() : new System.Collections.Generic.List<MaterialCategoryData>();
            foreach (MaterialCategoryData category in list.Where(x => string.IsNullOrWhiteSpace(x.ParentCode)).OrderBy(x => x.Name ?? x.Code))
            {
                tree.Items.Add(CreateCategoryTreeItem(category, list));
            }
        }

        private static TreeViewItem CreateCategoryTreeItem(MaterialCategoryData category, System.Collections.Generic.IList<MaterialCategoryData> allCategories)
        {
            var item = new TreeViewItem
            {
                Header = string.IsNullOrWhiteSpace(category.Name) ? category.Code : category.Name,
                Tag = category,
                IsExpanded = true
            };

            foreach (MaterialCategoryData child in allCategories.Where(x => string.Equals(x.ParentCode, category.Code, StringComparison.OrdinalIgnoreCase)).OrderBy(x => x.Name ?? x.Code))
            {
                item.Items.Add(CreateCategoryTreeItem(child, allCategories));
            }

            return item;
        }

        private bool FilterBoardMaterial(object value)
        {
            var material = value as BoardMaterialData;
            if (material == null)
            {
                return false;
            }

            if (string.IsNullOrWhiteSpace(selectedBoardCategoryCode))
            {
                return true;
            }

            return IsMaterialInCategoryTree(material.CategoryCode, selectedBoardCategoryCode, database.BoardCategories);
        }

        private bool FilterCoatingMaterial(object value)
        {
            var material = value as CoatingMaterialData;
            if (material == null)
            {
                return false;
            }

            if (string.IsNullOrWhiteSpace(selectedCoatingCategoryCode))
            {
                return true;
            }

            return IsMaterialInCategoryTree(material.CategoryCode, selectedCoatingCategoryCode, database.CoatingCategories);
        }

        private static bool IsMaterialInCategoryTree(string materialCategoryCode, string selectedCategoryCode, System.Collections.Generic.IList<MaterialCategoryData> categories)
        {
            if (string.IsNullOrWhiteSpace(selectedCategoryCode))
            {
                return true;
            }

            if (string.IsNullOrWhiteSpace(materialCategoryCode))
            {
                return false;
            }

            if (string.Equals(materialCategoryCode, selectedCategoryCode, StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }

            foreach (MaterialCategoryData category in categories ?? new System.Collections.Generic.List<MaterialCategoryData>())
            {
                if (string.Equals(category.ParentCode, selectedCategoryCode, StringComparison.OrdinalIgnoreCase)
                    && IsMaterialInCategoryTree(materialCategoryCode, category.Code, categories))
                {
                    return true;
                }
            }

            return false;
        }

        private void RefreshMaterialViews()
        {
            if (boardMaterialsView != null)
            {
                boardMaterialsView.Refresh();
            }

            if (coatingMaterialsView != null)
            {
                coatingMaterialsView.Refresh();
            }

            boardList.ItemsSource = null;
            boardList.ItemsSource = boardMaterialsView;
            coatingList.ItemsSource = null;
            coatingList.ItemsSource = coatingMaterialsView;
        }

        private void RefreshFormats(BoardMaterialData material)
        {
            formatList.ItemsSource = null;
            formatList.ItemsSource = material != null ? material.Formats : null;
            LoadSelectedFormat();
        }

        private void RefreshCategories()
        {
            updating = true;
            try
            {
                boardCategoryList.ItemsSource = null;
                boardCategoryList.ItemsSource = database.BoardCategories;
                coatingCategoryList.ItemsSource = null;
                coatingCategoryList.ItemsSource = database.CoatingCategories;
            }
            finally { updating = false; }

            LoadSelectedBoardCategory();
            LoadSelectedCoatingCategory();
        }

        private void SelectCoatingCombo(ComboBox combo, string code)
        {
            if (combo == null)
            {
                return;
            }

            var items = combo.ItemsSource as System.Collections.IEnumerable;
            if (items == null || string.IsNullOrWhiteSpace(code))
            {
                combo.SelectedItem = null;
                return;
            }

            foreach (object item in items)
            {
                CoatingMaterialData coating = item as CoatingMaterialData;
                if (coating != null && string.Equals(coating.Code, code, StringComparison.OrdinalIgnoreCase))
                {
                    combo.SelectedItem = coating;
                    return;
                }
            }

            combo.SelectedItem = null;
        }

        private static string SelectedCoatingCode(ComboBox combo)
        {
            var coating = combo != null ? combo.SelectedItem as CoatingMaterialData : null;
            return coating != null ? coating.Code : string.Empty;
        }

        private static string SelectedComboText(ComboBox combo, string fallback)
        {
            if (combo == null)
            {
                return fallback ?? string.Empty;
            }

            if (combo.SelectedItem is string text && !string.IsNullOrWhiteSpace(text))
            {
                return text;
            }

            if (combo.SelectedItem != null)
            {
                string selectedText = combo.SelectedItem.ToString();
                if (!string.IsNullOrWhiteSpace(selectedText))
                {
                    return selectedText;
                }
            }

            return fallback ?? string.Empty;
        }

        private static bool ConfirmDelete(string kind, string name)
        {
            return MessageBox.Show(
                "Удалить " + kind + " \"" + name + "\"?",
                "Подтверждение удаления",
                MessageBoxButton.YesNo,
                MessageBoxImage.Warning) == MessageBoxResult.Yes;
        }

        private void ApplyLayoutFromDatabase()
        {
            if (database == null || database.UiLayout == null)
            {
                return;
            }

            if (boardSidebarColumn != null)
            {
                boardSidebarColumn.Width = new GridLength(Math.Max(180.0, database.UiLayout.BoardSidebarWidth));
            }

            if (boardTreeRow != null)
            {
                boardTreeRow.Height = new GridLength(Math.Max(120.0, database.UiLayout.BoardTreeHeight));
            }

            if (coatingSidebarColumn != null)
            {
                coatingSidebarColumn.Width = new GridLength(Math.Max(180.0, database.UiLayout.CoatingSidebarWidth));
            }

            if (coatingTreeRow != null)
            {
                coatingTreeRow.Height = new GridLength(Math.Max(120.0, database.UiLayout.CoatingTreeHeight));
            }
        }

        private void CaptureLayoutToDatabase()
        {
            if (database == null)
            {
                return;
            }

            if (database.UiLayout == null)
            {
                database.UiLayout = new MaterialDatabaseUiLayout();
            }

            if (boardSidebarColumn != null)
            {
                database.UiLayout.BoardSidebarWidth = Math.Max(180.0, boardSidebarColumn.ActualWidth > 0.0 ? boardSidebarColumn.ActualWidth : boardSidebarColumn.Width.Value);
            }

            if (boardTreeRow != null)
            {
                database.UiLayout.BoardTreeHeight = Math.Max(120.0, boardTreeRow.ActualHeight > 0.0 ? boardTreeRow.ActualHeight : boardTreeRow.Height.Value);
            }

            if (coatingSidebarColumn != null)
            {
                database.UiLayout.CoatingSidebarWidth = Math.Max(180.0, coatingSidebarColumn.ActualWidth > 0.0 ? coatingSidebarColumn.ActualWidth : coatingSidebarColumn.Width.Value);
            }

            if (coatingTreeRow != null)
            {
                database.UiLayout.CoatingTreeHeight = Math.Max(120.0, coatingTreeRow.ActualHeight > 0.0 ? coatingTreeRow.ActualHeight : coatingTreeRow.Height.Value);
            }
        }

        private static Grid TwoColumnGrid(double leftWidth)
        {
            var grid = new Grid();
            grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(leftWidth) });
            grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
            return grid;
        }

        private static StackPanel Stack()
        {
            return new StackPanel { Margin = new Thickness(12, 0, 0, 0) };
        }

        private static TextBlock Label(string text)
        {
            return new TextBlock { Text = text, Foreground = Brushes.Gainsboro, Margin = new Thickness(0, 8, 0, 4) };
        }

        private static TextBox CreateTextBox()
        {
            return new TextBox
            {
                Margin = FieldMargin,
                Padding = new Thickness(8, 5, 8, 5),
                Background = new SolidColorBrush(Color.FromRgb(17, 24, 39)),
                Foreground = new SolidColorBrush(Color.FromRgb(248, 250, 252)),
                BorderBrush = new SolidColorBrush(Color.FromRgb(51, 65, 85)),
                BorderThickness = new Thickness(1)
            };
        }

        private static ComboBox CreateCombo(params string[] values)
        {
            var combo = new ComboBox
            {
                IsEditable = false,
                Margin = FieldMargin,
                Background = new SolidColorBrush(Color.FromRgb(17, 24, 39)),
                Foreground = new SolidColorBrush(Color.FromRgb(248, 250, 252)),
                BorderBrush = new SolidColorBrush(Color.FromRgb(51, 65, 85)),
                BorderThickness = new Thickness(1)
            };
            foreach (string value in values) combo.Items.Add(value);
            return combo;
        }

        private static ComboBox CreateMaterialComboBox()
        {
            return new ComboBox
            {
                IsEditable = false,
                Margin = FieldMargin,
                Background = new SolidColorBrush(Color.FromRgb(17, 24, 39)),
                Foreground = new SolidColorBrush(Color.FromRgb(248, 250, 252)),
                BorderBrush = new SolidColorBrush(Color.FromRgb(51, 65, 85)),
                BorderThickness = new Thickness(1),
                ItemTemplate = CreateMaterialItemTemplate()
            };
        }

        private static DataTemplate CreateMaterialItemTemplate()
        {
            var template = new DataTemplate();
            var text = new FrameworkElementFactory(typeof(TextBlock));
            text.SetBinding(TextBlock.TextProperty, new Binding("Name"));
            text.SetValue(TextBlock.ForegroundProperty, new SolidColorBrush(Color.FromRgb(226, 232, 240)));
            text.SetValue(TextBlock.VerticalAlignmentProperty, VerticalAlignment.Center);
            template.VisualTree = text;
            return template;
        }

        private static void ApplyButtonStyle(Button button)
        {
            button.Background = new SolidColorBrush(Color.FromRgb(37, 99, 235));
            button.Foreground = Brushes.White;
            button.BorderBrush = new SolidColorBrush(Color.FromRgb(96, 165, 250));
            button.BorderThickness = new Thickness(1);
            button.FontWeight = FontWeights.SemiBold;
        }

        private static Style CreateTextBoxStyle()
        {
            var style = new Style(typeof(TextBox));
            style.Setters.Add(new Setter(Control.BackgroundProperty, new SolidColorBrush(Color.FromRgb(17, 24, 39))));
            style.Setters.Add(new Setter(Control.ForegroundProperty, new SolidColorBrush(Color.FromRgb(248, 250, 252))));
            style.Setters.Add(new Setter(Control.BorderBrushProperty, new SolidColorBrush(Color.FromRgb(51, 65, 85))));
            style.Setters.Add(new Setter(Control.BorderThicknessProperty, new Thickness(1)));
            return style;
        }

        private static Style CreateComboStyle()
        {
            var style = new Style(typeof(ComboBox));
            style.Setters.Add(new Setter(Control.BackgroundProperty, new SolidColorBrush(Color.FromRgb(17, 24, 39))));
            style.Setters.Add(new Setter(Control.ForegroundProperty, new SolidColorBrush(Color.FromRgb(248, 250, 252))));
            style.Setters.Add(new Setter(Control.BorderBrushProperty, new SolidColorBrush(Color.FromRgb(51, 65, 85))));
            style.Setters.Add(new Setter(Control.BorderThicknessProperty, new Thickness(1)));
            return style;
        }

        private static Style CreateComboBoxItemStyle()
        {
            var style = new Style(typeof(ComboBoxItem));
            style.Setters.Add(new Setter(Control.BackgroundProperty, new SolidColorBrush(Color.FromRgb(11, 18, 32))));
            style.Setters.Add(new Setter(Control.ForegroundProperty, new SolidColorBrush(Color.FromRgb(226, 232, 240))));
            style.Setters.Add(new Setter(Control.PaddingProperty, new Thickness(8, 4, 8, 4)));
            style.Triggers.Add(new Trigger
            {
                Property = ComboBoxItem.IsSelectedProperty,
                Value = true,
                Setters =
                {
                    new Setter(Control.BackgroundProperty, new SolidColorBrush(Color.FromRgb(37, 99, 235))),
                    new Setter(Control.ForegroundProperty, Brushes.White)
                }
            });
            return style;
        }

        private static Style CreateListBoxStyle()
        {
            var style = new Style(typeof(ListBox));
            style.Setters.Add(new Setter(Control.BackgroundProperty, new SolidColorBrush(Color.FromRgb(11, 18, 32))));
            style.Setters.Add(new Setter(Control.ForegroundProperty, new SolidColorBrush(Color.FromRgb(248, 250, 252))));
            style.Setters.Add(new Setter(Control.BorderBrushProperty, new SolidColorBrush(Color.FromRgb(51, 65, 85))));
            style.Setters.Add(new Setter(Control.BorderThicknessProperty, new Thickness(1)));
            style.Setters.Add(new Setter(Control.PaddingProperty, new Thickness(2)));
            return style;
        }

        private static Style CreateListBoxItemStyle()
        {
            var style = new Style(typeof(ListBoxItem));
            style.Setters.Add(new Setter(Control.BackgroundProperty, new SolidColorBrush(Color.FromRgb(11, 18, 32))));
            style.Setters.Add(new Setter(Control.ForegroundProperty, new SolidColorBrush(Color.FromRgb(226, 232, 240))));
            style.Setters.Add(new Setter(Control.PaddingProperty, new Thickness(8, 5, 8, 5)));
            style.Setters.Add(new Setter(Control.MarginProperty, new Thickness(0, 0, 0, 1)));
            style.Setters.Add(new Setter(Control.HorizontalContentAlignmentProperty, HorizontalAlignment.Stretch));
            style.Triggers.Add(new Trigger
            {
                Property = ListBoxItem.IsSelectedProperty,
                Value = true,
                Setters =
                {
                    new Setter(Control.BackgroundProperty, new SolidColorBrush(Color.FromRgb(37, 99, 235))),
                    new Setter(Control.ForegroundProperty, Brushes.White)
                }
            });
            return style;
        }

        private static Style CreateTreeViewStyle()
        {
            var style = new Style(typeof(TreeView));
            style.Setters.Add(new Setter(Control.BackgroundProperty, new SolidColorBrush(Color.FromRgb(11, 18, 32))));
            style.Setters.Add(new Setter(Control.ForegroundProperty, new SolidColorBrush(Color.FromRgb(248, 250, 252))));
            style.Setters.Add(new Setter(Control.BorderBrushProperty, new SolidColorBrush(Color.FromRgb(51, 65, 85))));
            style.Setters.Add(new Setter(Control.BorderThicknessProperty, new Thickness(1)));
            style.Setters.Add(new Setter(Control.PaddingProperty, new Thickness(4)));
            return style;
        }

        private static Style CreateTreeViewItemStyle()
        {
            var style = new Style(typeof(TreeViewItem));
            style.Setters.Add(new Setter(Control.BackgroundProperty, new SolidColorBrush(Color.FromRgb(11, 18, 32))));
            style.Setters.Add(new Setter(Control.ForegroundProperty, new SolidColorBrush(Color.FromRgb(226, 232, 240))));
            style.Setters.Add(new Setter(Control.PaddingProperty, new Thickness(4, 3, 4, 3)));
            style.Setters.Add(new Setter(Control.MarginProperty, new Thickness(0, 0, 0, 1)));
            style.Setters.Add(new Setter(TreeViewItem.IsExpandedProperty, true));
            style.Setters.Add(new Setter(Control.BorderBrushProperty, new SolidColorBrush(Color.FromRgb(51, 65, 85))));
            style.Setters.Add(new Setter(Control.BorderThicknessProperty, new Thickness(1)));
            style.Setters.Add(new Setter(Control.TemplateProperty, CreateTreeViewItemTemplate()));
            style.Triggers.Add(new Trigger
            {
                Property = TreeViewItem.IsSelectedProperty,
                Value = true,
                Setters =
                {
                    new Setter(Control.BackgroundProperty, new SolidColorBrush(Color.FromRgb(37, 99, 235))),
                    new Setter(Control.ForegroundProperty, Brushes.White)
                }
            });
            style.Triggers.Add(new Trigger
            {
                Property = Selector.IsSelectionActiveProperty,
                Value = false,
                Setters =
                {
                    new Setter(Control.BackgroundProperty, new SolidColorBrush(Color.FromRgb(30, 41, 59))),
                    new Setter(Control.ForegroundProperty, new SolidColorBrush(Color.FromRgb(226, 232, 240)))
                }
            });
            return style;
        }

        private static ControlTemplate CreateTreeViewItemTemplate()
        {
            var template = new ControlTemplate(typeof(TreeViewItem));
            var dock = new FrameworkElementFactory(typeof(DockPanel));
            dock.SetValue(DockPanel.LastChildFillProperty, true);

            var headerBorder = new FrameworkElementFactory(typeof(Border));
            headerBorder.SetValue(DockPanel.DockProperty, Dock.Top);
            headerBorder.SetBinding(Border.BackgroundProperty, new Binding("Background") { RelativeSource = new RelativeSource(RelativeSourceMode.TemplatedParent) });
            headerBorder.SetBinding(Border.BorderBrushProperty, new Binding("BorderBrush") { RelativeSource = new RelativeSource(RelativeSourceMode.TemplatedParent) });
            headerBorder.SetBinding(Border.BorderThicknessProperty, new Binding("BorderThickness") { RelativeSource = new RelativeSource(RelativeSourceMode.TemplatedParent) });
            headerBorder.SetBinding(Border.PaddingProperty, new Binding("Padding") { RelativeSource = new RelativeSource(RelativeSourceMode.TemplatedParent) });
            headerBorder.SetValue(Border.SnapsToDevicePixelsProperty, true);

            var header = new FrameworkElementFactory(typeof(ContentPresenter));
            header.SetValue(ContentPresenter.ContentSourceProperty, "Header");
            header.SetValue(ContentPresenter.VerticalAlignmentProperty, VerticalAlignment.Center);
            header.SetValue(ContentPresenter.MarginProperty, new Thickness(4, 1, 4, 1));
            headerBorder.AppendChild(header);
            dock.AppendChild(headerBorder);

            var items = new FrameworkElementFactory(typeof(ItemsPresenter));
            items.SetValue(FrameworkElement.MarginProperty, new Thickness(14, 2, 0, 0));
            dock.AppendChild(items);

            template.VisualTree = dock;
            return template;
        }

        private static Style CreateTabControlStyle()
        {
            var style = new Style(typeof(TabControl));
            style.Setters.Add(new Setter(Control.BackgroundProperty, new SolidColorBrush(Color.FromRgb(15, 23, 42))));
            style.Setters.Add(new Setter(Control.BorderThicknessProperty, new Thickness(0)));
            style.Setters.Add(new Setter(Control.TemplateProperty, CreateTabControlTemplate()));
            return style;
        }

        private static Style CreateTabItemStyle()
        {
            var style = new Style(typeof(TabItem));
            style.Setters.Add(new Setter(Control.ForegroundProperty, new SolidColorBrush(Color.FromRgb(226, 232, 240))));
            style.Setters.Add(new Setter(Control.BackgroundProperty, new SolidColorBrush(Color.FromRgb(17, 24, 39))));
            style.Setters.Add(new Setter(Control.PaddingProperty, new Thickness(12, 6, 12, 6)));
            style.Setters.Add(new Setter(Control.TemplateProperty, CreateTabItemTemplate()));
            style.Triggers.Add(new Trigger
            {
                Property = TabItem.IsSelectedProperty,
                Value = true,
                Setters =
                {
                    new Setter(Control.BackgroundProperty, new SolidColorBrush(Color.FromRgb(37, 99, 235))),
                    new Setter(Control.ForegroundProperty, Brushes.White)
                }
            });
            return style;
        }

        private static ControlTemplate CreateTabControlTemplate()
        {
            var template = new ControlTemplate(typeof(TabControl));
            var dock = new FrameworkElementFactory(typeof(DockPanel));
            dock.SetValue(DockPanel.LastChildFillProperty, true);

            var tabPanel = new FrameworkElementFactory(typeof(TabPanel));
            tabPanel.SetValue(DockPanel.DockProperty, Dock.Top);
            tabPanel.SetValue(Panel.IsItemsHostProperty, true);
            tabPanel.SetValue(FrameworkElement.MarginProperty, new Thickness(0, 0, 0, 8));
            dock.AppendChild(tabPanel);

            var border = new FrameworkElementFactory(typeof(Border));
            border.SetBinding(Border.BackgroundProperty, new Binding("Background") { RelativeSource = new RelativeSource(RelativeSourceMode.TemplatedParent) });
            border.SetValue(Border.BorderBrushProperty, new SolidColorBrush(Color.FromRgb(51, 65, 85)));
            border.SetValue(Border.BorderThicknessProperty, new Thickness(0));
            border.SetValue(Border.SnapsToDevicePixelsProperty, true);

            var presenter = new FrameworkElementFactory(typeof(ContentPresenter));
            presenter.SetValue(ContentPresenter.ContentSourceProperty, "SelectedContent");
            presenter.SetValue(ContentPresenter.MarginProperty, new Thickness(0));
            border.AppendChild(presenter);
            dock.AppendChild(border);

            template.VisualTree = dock;
            return template;
        }

        private static ControlTemplate CreateTabItemTemplate()
        {
            var template = new ControlTemplate(typeof(TabItem));
            var border = new FrameworkElementFactory(typeof(Border));
            border.Name = "BorderRoot";
            border.SetBinding(Border.BackgroundProperty, new Binding("Background") { RelativeSource = new RelativeSource(RelativeSourceMode.TemplatedParent) });
            border.SetBinding(Border.BorderBrushProperty, new Binding("BorderBrush") { RelativeSource = new RelativeSource(RelativeSourceMode.TemplatedParent) });
            border.SetBinding(Border.BorderThicknessProperty, new Binding("BorderThickness") { RelativeSource = new RelativeSource(RelativeSourceMode.TemplatedParent) });
            border.SetValue(Border.SnapsToDevicePixelsProperty, true);

            var content = new FrameworkElementFactory(typeof(ContentPresenter));
            content.SetValue(ContentPresenter.ContentSourceProperty, "Header");
            content.SetValue(ContentPresenter.HorizontalAlignmentProperty, HorizontalAlignment.Center);
            content.SetValue(ContentPresenter.VerticalAlignmentProperty, VerticalAlignment.Center);
            content.SetValue(ContentPresenter.MarginProperty, new Thickness(12, 5, 12, 5));
            border.AppendChild(content);

            template.VisualTree = border;
            template.Triggers.Add(new Trigger
            {
                Property = TabItem.IsSelectedProperty,
                Value = true,
                Setters =
                {
                    new Setter(Control.BackgroundProperty, new SolidColorBrush(Color.FromRgb(37, 99, 235))),
                    new Setter(Control.ForegroundProperty, Brushes.White)
                }
            });
            return template;
        }

        private static Thickness FieldMargin => new Thickness(0, 0, 0, 4);

        private static UIElement WrapList(ListBox list, string title, Action add, Action delete)
        {
            var dock = new DockPanel { Margin = new Thickness(0, 0, 12, 0) };
            var top = new StackPanel { Orientation = Orientation.Horizontal, Margin = new Thickness(0, 0, 0, 8) };
            top.Children.Add(new TextBlock { Text = title, Foreground = Brushes.White, FontWeight = FontWeights.SemiBold, VerticalAlignment = VerticalAlignment.Center });
            var addButton = new Button { Content = "+", Width = 28, Margin = new Thickness(8, 0, 4, 0) };
            var delButton = new Button { Content = "-", Width = 28 };
            addButton.Background = new SolidColorBrush(Color.FromRgb(37, 99, 235));
            addButton.Foreground = Brushes.White;
            addButton.BorderBrush = new SolidColorBrush(Color.FromRgb(96, 165, 250));
            delButton.Background = new SolidColorBrush(Color.FromRgb(37, 99, 235));
            delButton.Foreground = Brushes.White;
            delButton.BorderBrush = new SolidColorBrush(Color.FromRgb(96, 165, 250));
            addButton.Click += delegate { add(); };
            delButton.Click += delegate { delete(); };
            top.Children.Add(addButton);
            top.Children.Add(delButton);
            DockPanel.SetDock(top, Dock.Top);
            dock.Children.Add(top);
            dock.Children.Add(list);
            return dock;
        }

        private static bool TryDouble(string text, out double value)
        {
            return double.TryParse(text, NumberStyles.Float, CultureInfo.InvariantCulture, out value)
                || double.TryParse(text, NumberStyles.Float, CultureInfo.CurrentCulture, out value);
        }

        private static string UniqueCode(string prefix, System.Collections.Generic.IEnumerable<string> existing)
        {
            int index = 1;
            string code;
            var set = existing == null
                ? new System.Collections.Generic.HashSet<string>(StringComparer.OrdinalIgnoreCase)
                : new System.Collections.Generic.HashSet<string>(existing.Where(x => !string.IsNullOrWhiteSpace(x)), StringComparer.OrdinalIgnoreCase);
            do
            {
                code = prefix + "-" + index.ToString(CultureInfo.InvariantCulture);
                index++;
            }
            while (set.Contains(code));
            return code;
        }
    }
}
