using System;
using System.Globalization;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Shapes;
using AutoCAD_BoardSorter.Models;
using AcadApp = Autodesk.AutoCAD.ApplicationServices.Application;

namespace AutoCAD_BoardSorter.Ui
{
    public partial class AssemblyEditorWindow : Window
    {
        private readonly AssemblyEditorController controller;
        private MaterialDatabase materialDatabase;
        private bool isPanning;
        private Point lastPanPoint;
        private bool updatingDrawerUi;

        internal AssemblyEditorWindow(IAssemblyEditorBackend backend)
        {
            InitializeComponent();
            controller = new AssemblyEditorController(backend);
            LoadMaterialChoices();
            MaterialComboBox.Text = controller.State.ActiveMaterial;
            ThicknessTextBox.Text = controller.State.ActiveThickness.ToString("0.#", CultureInfo.InvariantCulture);
            OffsetTextBox.Text = controller.State.ActiveOffset.ToString("0.#", CultureInfo.InvariantCulture);
            SyncDrawerSettingsToUi();
            SelectToolButton.IsChecked = true;
        }

        internal void LoadScene(AssemblyScene scene)
        {
            controller.LoadScene(scene);
            Render();
        }

        private void OnToolButtonClick(object sender, RoutedEventArgs e)
        {
            ToggleButton clicked = sender as ToggleButton;
            if (clicked == null)
            {
                return;
            }

            SelectToolButton.IsChecked = ReferenceEquals(clicked, SelectToolButton);
            VerticalToolButton.IsChecked = ReferenceEquals(clicked, VerticalToolButton);
            HorizontalToolButton.IsChecked = ReferenceEquals(clicked, HorizontalToolButton);
            FrontToolButton.IsChecked = ReferenceEquals(clicked, FrontToolButton);
            BackToolButton.IsChecked = ReferenceEquals(clicked, BackToolButton);
            DrawersToolButton.IsChecked = ReferenceEquals(clicked, DrawersToolButton);

            if (ReferenceEquals(clicked, VerticalToolButton))
            {
                controller.SetTool(AssemblyEditorTool.VerticalPanel);
            }
            else if (ReferenceEquals(clicked, HorizontalToolButton))
            {
                controller.SetTool(AssemblyEditorTool.HorizontalPanel);
            }
            else if (ReferenceEquals(clicked, FrontToolButton))
            {
                controller.SetTool(AssemblyEditorTool.FrontPanel);
            }
            else if (ReferenceEquals(clicked, BackToolButton))
            {
                controller.SetTool(AssemblyEditorTool.BackPanel);
            }
            else if (ReferenceEquals(clicked, DrawersToolButton))
            {
                controller.SetTool(AssemblyEditorTool.Drawers);
            }
            else
            {
                controller.SetTool(AssemblyEditorTool.Select);
            }

            DrawerSettingsPanel.Visibility = controller.State.Tool == AssemblyEditorTool.Drawers ? Visibility.Visible : Visibility.Collapsed;
            Render();
        }

        private void OnDrawerModeChanged(object sender, RoutedEventArgs e)
        {
            if (updatingDrawerUi)
            {
                return;
            }

            controller.UpdateDrawerSettings(settings =>
            {
                settings.Mode = DrawerInsetRadio.IsChecked == true ? AssemblyDrawerMode.Inset : AssemblyDrawerMode.Overlay;
            });
            Render();
        }

        private void OnDrawerSettingsChanged(object sender, RoutedEventArgs e)
        {
            if (updatingDrawerUi)
            {
                return;
            }

            controller.UpdateDrawerSettings(settings =>
            {
                if (int.TryParse(DrawerCountTextBox.Text, NumberStyles.Integer, CultureInfo.InvariantCulture, out int count)
                    || int.TryParse(DrawerCountTextBox.Text, NumberStyles.Integer, CultureInfo.CurrentCulture, out count))
                {
                    settings.Count = Math.Max(1, count);
                }

                settings.Mode = DrawerInsetRadio.IsChecked == true ? AssemblyDrawerMode.Inset : AssemblyDrawerMode.Overlay;
                if (TryParseDouble(DrawerGapLeftTextBox.Text, out double gapLeft)) settings.GapLeft = Math.Abs(gapLeft);
                if (TryParseDouble(DrawerGapRightTextBox.Text, out double gapRight)) settings.GapRight = Math.Abs(gapRight);
                if (TryParseDouble(DrawerGapTopTextBox.Text, out double gapTop)) settings.GapTop = Math.Abs(gapTop);
                if (TryParseDouble(DrawerGapBottomTextBox.Text, out double gapBottom)) settings.GapBottom = Math.Abs(gapBottom);
                if (TryParseDouble(DrawerGapBetweenTextBox.Text, out double gapBetween)) settings.GapBetween = Math.Max(0.0, gapBetween);
                if (TryParseDouble(DrawerFrontGapTextBox.Text, out double frontGap)) settings.FrontGap = Math.Max(0.0, frontGap);
                if (TryParseDouble(DrawerFrontThicknessTextBox.Text, out double frontThickness)) settings.FrontThickness = Math.Max(1.0, frontThickness);
                if (TryParseDouble(DrawerBottomThicknessTextBox.Text, out double bottomThickness)) settings.BottomThickness = Math.Max(1.0, bottomThickness);
                if (TryParseDouble(DrawerDepthTextBox.Text, out double depth)) settings.Depth = Math.Max(1.0, depth);
                settings.AutoDepth = DrawerAutoDepthCheckBox.IsChecked == true;
                settings.FrontMaterial = DrawerFrontMaterialComboBox.Text;
                settings.BodyMaterial = DrawerBodyMaterialComboBox.Text;
                settings.BottomMaterial = DrawerBottomMaterialComboBox.Text;
                double frontMaterialThickness = ResolveMaterialThickness(settings.FrontMaterial);
                if (frontMaterialThickness > 0.0 && ReferenceEquals(sender, DrawerFrontMaterialComboBox))
                {
                    settings.FrontThickness = frontMaterialThickness;
                }

                double bottomMaterialThickness = ResolveMaterialThickness(settings.BottomMaterial);
                if (bottomMaterialThickness > 0.0 && ReferenceEquals(sender, DrawerBottomMaterialComboBox))
                {
                    settings.BottomThickness = bottomMaterialThickness;
                }

                ApplyDrawerHeights(settings);
            });

            if (ReferenceEquals(sender, DrawerCountTextBox)
                || ReferenceEquals(sender, DrawerFrontMaterialComboBox)
                || ReferenceEquals(sender, DrawerBottomMaterialComboBox))
            {
                SyncDrawerSettingsToUi();
            }

            Render();
        }

        private void OnDrawerSettingsChanged(object sender, TextChangedEventArgs e)
        {
            OnDrawerSettingsChanged(sender, (RoutedEventArgs)e);
        }

        private void OnDrawerSettingsChanged(object sender, SelectionChangedEventArgs e)
        {
            OnDrawerSettingsChanged(sender, new RoutedEventArgs());
        }

        private void OnDrawerHeightAutoClick(object sender, RoutedEventArgs e)
        {
            Button button = sender as Button;
            if (button == null || !(button.Tag is int))
            {
                return;
            }

            int index = (int)button.Tag;
            controller.UpdateDrawerSettings(settings =>
            {
                AssemblyDrawerLayoutCalculator.EnsureDrafts(settings);
                if (index >= 0 && index < settings.Drawers.Count)
                {
                    settings.Drawers[index].AutoHeight = true;
                    settings.Drawers[index].Height = 0.0;
                }
            });
            SyncDrawerSettingsToUi();
            Render();
        }

        private void OnDrawerComboLostFocus(object sender, KeyboardFocusChangedEventArgs e)
        {
            OnDrawerSettingsChanged(sender, new RoutedEventArgs());
        }

        private void OnMaterialComboLostFocus(object sender, KeyboardFocusChangedEventArgs e)
        {
            ApplyActiveMaterialFromCombo(true);
        }

        private void OnMaterialChanged(object sender, SelectionChangedEventArgs e)
        {
            ApplyActiveMaterialFromCombo(true);
            RenderSummary();
        }

        private void OnThicknessChanged(object sender, TextChangedEventArgs e)
        {
            if (TryParseDouble(ThicknessTextBox.Text, out double thickness))
            {
                controller.SetThickness(thickness);
                Render();
            }
        }

        private void OnOffsetChanged(object sender, TextChangedEventArgs e)
        {
            if (TryParseDouble(OffsetTextBox.Text, out double offset))
            {
                controller.SetOffset(offset);
                Render();
            }
        }

        private void OnReloadClick(object sender, RoutedEventArgs e)
        {
            LoadMaterialChoices();
            controller.Reload();
            Render();
        }

        private void OnCommitClick(object sender, RoutedEventArgs e)
        {
            try
            {
                var db = AcadApp.DocumentManager.MdiActiveDocument != null
                    ? AcadApp.DocumentManager.MdiActiveDocument.Database
                    : null;
                PaletteDebugLogger.Info(db, "AssemblyEditorWindow OnCommitClick");
                if (controller.CommitPreview())
                {
                    Render();
                }
            }
            catch (Exception ex)
            {
                var doc = AcadApp.DocumentManager.MdiActiveDocument;
                PaletteDebugLogger.Error(doc != null ? doc.Database : null, "AssemblyEditorWindow OnCommitClick failed", ex);
                if (doc != null)
                {
                    doc.Editor.WriteMessage("\nОшибка конструктора сборок: {0}", ex.Message);
                }
            }
        }

        private void OnViewportSizeChanged(object sender, SizeChangedEventArgs e)
        {
            controller.SetViewportSize(e.NewSize);
            Render();
        }

        private void OnViewportMouseMove(object sender, System.Windows.Input.MouseEventArgs e)
        {
            if (isPanning)
            {
                Point current = e.GetPosition(ViewportCanvas);
                controller.Pan(current.X - lastPanPoint.X, current.Y - lastPanPoint.Y);
                lastPanPoint = current;
                Render();
                return;
            }

            controller.HandlePointerMove(e.GetPosition(ViewportCanvas));
            Render();
        }

        private void OnViewportMouseLeave(object sender, System.Windows.Input.MouseEventArgs e)
        {
            controller.HandlePointerLeave();
            Render();
        }

        private void OnViewportMouseLeftButtonDown(object sender, System.Windows.Input.MouseButtonEventArgs e)
        {
            try
            {
                var doc = AcadApp.DocumentManager.MdiActiveDocument;
                var db = doc != null ? doc.Database : null;
                Point point = e.GetPosition(ViewportCanvas);
                PaletteDebugLogger.Info(
                    db,
                    "AssemblyEditorWindow MouseLeftButtonDown tool=" + controller.State.Tool
                    + " point=" + point.X.ToString("0.##", CultureInfo.InvariantCulture)
                    + "," + point.Y.ToString("0.##", CultureInfo.InvariantCulture));

                bool hit = controller.HandleClick(point);
                if (controller.State.Tool != AssemblyEditorTool.Select && hit)
                {
                    bool committed = controller.CommitPreview();
                    PaletteDebugLogger.Info(db, "AssemblyEditorWindow MouseLeftButtonDown committed=" + committed);
                }
                else
                {
                    PaletteDebugLogger.Info(db, "AssemblyEditorWindow MouseLeftButtonDown commit skipped hit=" + hit);
                }

                Render();
            }
            catch (Exception ex)
            {
                var doc = AcadApp.DocumentManager.MdiActiveDocument;
                PaletteDebugLogger.Error(doc != null ? doc.Database : null, "AssemblyEditorWindow MouseLeftButtonDown failed", ex);
                if (doc != null)
                {
                    doc.Editor.WriteMessage("\nОшибка обработки клика в конструкторе: {0}", ex.Message);
                }
            }
        }

        private void OnViewportMouseRightButtonDown(object sender, MouseButtonEventArgs e)
        {
            isPanning = true;
            lastPanPoint = e.GetPosition(ViewportCanvas);
            ViewportCanvas.CaptureMouse();
        }

        private void OnViewportMouseRightButtonUp(object sender, MouseButtonEventArgs e)
        {
            isPanning = false;
            if (ViewportCanvas.IsMouseCaptured)
            {
                ViewportCanvas.ReleaseMouseCapture();
            }
        }

        private void OnViewportMouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
        }

        private void OnViewportMouseWheel(object sender, MouseWheelEventArgs e)
        {
            controller.ZoomAt(e.GetPosition(ViewportCanvas), e.Delta > 0 ? 1.1 : (1.0 / 1.1));
            Render();
        }

        private void OnResetViewClick(object sender, RoutedEventArgs e)
        {
            controller.ResetView();
            Render();
        }

        private void OnUndoClick(object sender, RoutedEventArgs e)
        {
            ExecuteAcadCommand("_.UNDO 1 ");
        }

        private void OnRedoClick(object sender, RoutedEventArgs e)
        {
            ExecuteAcadCommand("_.REDO ");
        }

        private void OnWindowPreviewKeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.LeftShift || e.Key == Key.RightShift)
            {
                controller.SetShiftPressed(true);
                Render();
            }
        }

        private void OnWindowPreviewKeyUp(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.LeftShift || e.Key == Key.RightShift)
            {
                controller.SetShiftPressed(false);
                Render();
            }
        }

        private void Render()
        {
            ViewportCanvas.Children.Clear();
            AssemblyEditorRenderModel render = controller.RenderModel;
            if (render == null || render.Projection == null)
            {
                RenderSummary();
                return;
            }

            Rectangle containerRect = new Rectangle
            {
                Width = Math.Max(1.0, render.ContainerRect.Width),
                Height = Math.Max(1.0, render.ContainerRect.Height),
                Stroke = new SolidColorBrush(Color.FromRgb(148, 163, 184)),
                StrokeThickness = 1.2,
                RadiusX = 3,
                RadiusY = 3,
                IsHitTestVisible = false
            };
            Canvas.SetLeft(containerRect, render.ContainerRect.Left);
            Canvas.SetTop(containerRect, render.ContainerRect.Top);
            ViewportCanvas.Children.Add(containerRect);

            foreach (AssemblyEditorNicheVisual niche in render.Niches)
            {
                Rectangle rect = new Rectangle
                {
                    Width = Math.Max(1.0, niche.Bounds.Width),
                    Height = Math.Max(1.0, niche.Bounds.Height),
                    Fill = niche.Fill,
                    Stroke = niche.Stroke,
                    StrokeThickness = niche.IsSelected ? 2.0 : 1.0,
                    RadiusX = 0,
                    RadiusY = 0,
                    IsHitTestVisible = false
                };
                Canvas.SetLeft(rect, niche.Bounds.Left);
                Canvas.SetTop(rect, niche.Bounds.Top);
                ViewportCanvas.Children.Add(rect);

                AddCenteredLabel(niche.Bounds, niche.Label, niche.IsSelected ? Brushes.White : new SolidColorBrush(Color.FromRgb(203, 213, 225)), 12, FontWeights.SemiBold);
            }

            foreach (AssemblyEditorPanelVisual panel in render.Panels)
            {
                Rectangle rect = new Rectangle
                {
                    Width = Math.Max(1.0, panel.Bounds.Width),
                    Height = Math.Max(1.0, panel.Bounds.Height),
                    Fill = panel.Fill,
                    Stroke = panel.Stroke,
                    StrokeThickness = 1.0,
                    RadiusX = 3,
                    RadiusY = 3,
                    IsHitTestVisible = false
                };
                Canvas.SetLeft(rect, panel.Bounds.Left);
                Canvas.SetTop(rect, panel.Bounds.Top);
                ViewportCanvas.Children.Add(rect);

                if (panel.Bounds.Width > 72.0 && panel.Bounds.Height > 22.0)
                {
                    AddCenteredLabel(panel.Bounds, panel.Label, Brushes.White, 11, FontWeights.SemiBold);
                }
            }

            if (render.Preview != null)
            {
                foreach (Rect previewRect in render.Preview.ModelRects != null && render.Preview.ModelRects.Count > 0
                    ? render.Preview.ModelRects.Select(r => render.Projection.ModelRectToScreen(r.Left, r.Top, r.Right, r.Bottom))
                    : new[] { render.Preview.Bounds })
                {
                    Rectangle rect = new Rectangle
                    {
                        Width = Math.Max(1.0, previewRect.Width),
                        Height = Math.Max(1.0, previewRect.Height),
                        Fill = render.Preview.Fill,
                        Stroke = render.Preview.Stroke,
                        StrokeThickness = 2.0,
                        StrokeDashArray = new DoubleCollection(new[] { 6.0, 4.0 }),
                        RadiusX = 0,
                        RadiusY = 0,
                        IsHitTestVisible = false
                    };
                    Canvas.SetLeft(rect, previewRect.Left);
                    Canvas.SetTop(rect, previewRect.Top);
                    ViewportCanvas.Children.Add(rect);
                }

                AddPreviewLabel(render.Preview.Bounds, render.Preview.Label);
            }

            RenderSummary();
        }

        private void RenderSummary()
        {
            AssemblyEditorState state = controller.State;
            AssemblyScene scene = state.Scene;
            AssemblyNumberTextBlock.Text = scene != null && scene.Container != null ? Safe(scene.Container.AssemblyNumber) : "-";
            SelectedNicheTextBlock.Text = state.SelectedNiche != null
                ? string.Format(
                    CultureInfo.InvariantCulture,
                    "{0:0.#} x {1:0.#} x {2:0.#}\nID: {3}",
                    state.SelectedNiche.Width,
                    state.SelectedNiche.Height,
                    state.SelectedNiche.Depth,
                    Safe(state.SelectedNiche.Id))
                : "Ниша не выбрана";
            PreviewTextBlock.Text = controller.RenderModel != null && controller.RenderModel.Preview != null
                ? controller.RenderModel.Preview.Label.Replace("\n", Environment.NewLine)
                : "Preview появится при выборе инструмента и наведении на нишу.";
            PanelSummaryTextBlock.Text = scene == null
                ? "-"
                : string.Format(
                    CultureInfo.InvariantCulture,
                    "Панелей: {0}\nНиш: {1}\nИгнорируется тел: {2}",
                    scene.Panels.Count,
                    scene.Niches.Count,
                    scene.IgnoredSolids.Count);
            WarningsListBox.ItemsSource = scene != null ? scene.Warnings.ToList() : null;
            StatusTextBlock.Text = Safe(state.StatusText);
            SceneSummaryTextBlock.Text = scene == null || scene.Container == null
                ? string.Empty
                : string.Format(
                    CultureInfo.InvariantCulture,
                    "{0:0.#} x {1:0.#} x {2:0.#} мм",
                    scene.Container.Width,
                    scene.Container.Height,
                    scene.Container.Depth);
            CommitButton.IsEnabled = controller.RenderModel != null && controller.RenderModel.Preview != null && state.Tool != AssemblyEditorTool.Select;
            CommitButton.Visibility = state.Tool == AssemblyEditorTool.Select ? Visibility.Collapsed : Visibility.Visible;
            DrawerSettingsPanel.Visibility = state.Tool == AssemblyEditorTool.Drawers ? Visibility.Visible : Visibility.Collapsed;
        }

        private void AddCenteredLabel(Rect bounds, string text, Brush foreground, double fontSize, FontWeight fontWeight)
        {
            if (string.IsNullOrWhiteSpace(text))
            {
                return;
            }

            TextBlock label = new TextBlock
            {
                Text = text,
                Foreground = foreground,
                FontSize = fontSize,
                FontWeight = fontWeight,
                TextAlignment = TextAlignment.Center,
                TextWrapping = TextWrapping.Wrap,
                Width = Math.Max(36.0, bounds.Width - 10.0),
                IsHitTestVisible = false
            };

            label.Measure(new Size(label.Width, double.PositiveInfinity));
            Canvas.SetLeft(label, bounds.Left + ((bounds.Width - label.Width) * 0.5));
            Canvas.SetTop(label, bounds.Top + ((bounds.Height - label.DesiredSize.Height) * 0.5));
            ViewportCanvas.Children.Add(label);
        }

        private void AddPreviewLabel(Rect bounds, string text)
        {
            if (string.IsNullOrWhiteSpace(text))
            {
                return;
            }

            Border bubble = new Border
            {
                Background = new SolidColorBrush(Color.FromArgb(235, 15, 23, 42)),
                BorderBrush = new SolidColorBrush(Color.FromRgb(244, 114, 182)),
                BorderThickness = new Thickness(1),
                CornerRadius = new CornerRadius(6),
                Padding = new Thickness(8, 6, 8, 6),
                MaxWidth = 220,
                IsHitTestVisible = false
            };

            TextBlock label = new TextBlock
            {
                Text = text,
                Foreground = Brushes.White,
                FontSize = 11,
                FontWeight = FontWeights.SemiBold,
                TextWrapping = TextWrapping.Wrap,
                IsHitTestVisible = false
            };

            bubble.Child = label;
            bubble.Measure(new Size(220, double.PositiveInfinity));
            double left = Math.Max(8.0, Math.Min(ViewportCanvas.ActualWidth - bubble.DesiredSize.Width - 8.0, bounds.Left));
            double top = Math.Max(8.0, bounds.Top - bubble.DesiredSize.Height - 10.0);
            Canvas.SetLeft(bubble, left);
            Canvas.SetTop(bubble, top);
            ViewportCanvas.Children.Add(bubble);
        }

        private static bool TryParseDouble(string text, out double value)
        {
            return double.TryParse(text, NumberStyles.Float, CultureInfo.InvariantCulture, out value)
                || double.TryParse(text, NumberStyles.Float, CultureInfo.CurrentCulture, out value);
        }

        private void SyncDrawerSettingsToUi()
        {
            AssemblyDrawerSettings settings = controller.State.DrawerSettings;
            if (settings == null)
            {
                return;
            }

            updatingDrawerUi = true;
            try
            {
                AssemblyDrawerLayoutCalculator.EnsureDrafts(settings);
                DrawerCountTextBox.Text = settings.Count.ToString(CultureInfo.InvariantCulture);
                DrawerOverlayRadio.IsChecked = settings.Mode == AssemblyDrawerMode.Overlay;
                DrawerInsetRadio.IsChecked = settings.Mode == AssemblyDrawerMode.Inset;
                DrawerGapLeftTextBox.Text = settings.GapLeft.ToString("0.###", CultureInfo.InvariantCulture);
                DrawerGapRightTextBox.Text = settings.GapRight.ToString("0.###", CultureInfo.InvariantCulture);
                DrawerGapTopTextBox.Text = settings.GapTop.ToString("0.###", CultureInfo.InvariantCulture);
                DrawerGapBottomTextBox.Text = settings.GapBottom.ToString("0.###", CultureInfo.InvariantCulture);
                DrawerGapBetweenTextBox.Text = settings.GapBetween.ToString("0.###", CultureInfo.InvariantCulture);
                DrawerFrontGapTextBox.Text = settings.FrontGap.ToString("0.###", CultureInfo.InvariantCulture);
                DrawerFrontThicknessTextBox.Text = settings.FrontThickness.ToString("0.###", CultureInfo.InvariantCulture);
                DrawerBottomThicknessTextBox.Text = settings.BottomThickness.ToString("0.###", CultureInfo.InvariantCulture);
                DrawerAutoDepthCheckBox.IsChecked = settings.AutoDepth;
                DrawerDepthTextBox.Text = settings.Depth.ToString("0.###", CultureInfo.InvariantCulture);
                DrawerFrontMaterialComboBox.Text = settings.FrontMaterial ?? string.Empty;
                DrawerBodyMaterialComboBox.Text = settings.BodyMaterial ?? string.Empty;
                DrawerBottomMaterialComboBox.Text = settings.BottomMaterial ?? string.Empty;
                RebuildDrawerHeightRows(settings);
            }
            finally
            {
                updatingDrawerUi = false;
            }
        }

        private void ApplyDrawerHeights(AssemblyDrawerSettings settings)
        {
            AssemblyDrawerLayoutCalculator.EnsureDrafts(settings);
            foreach (UIElement child in DrawerHeightsPanel.Children)
            {
                Grid row = child as Grid;
                if (row == null || !(row.Tag is int))
                {
                    continue;
                }

                int index = (int)row.Tag;
                if (index < 0 || index >= settings.Drawers.Count)
                {
                    continue;
                }

                TextBox textBox = row.Children.OfType<TextBox>().FirstOrDefault();
                if (textBox == null)
                {
                    continue;
                }

                string text = (textBox.Text ?? string.Empty).Trim();
                if (string.Equals(text, "auto", StringComparison.OrdinalIgnoreCase)
                    || string.Equals(text, "авто", StringComparison.OrdinalIgnoreCase))
                {
                    settings.Drawers[index].AutoHeight = true;
                    settings.Drawers[index].Height = 0.0;
                    continue;
                }

                if (TryParseDouble(text, out double height))
                {
                    settings.Drawers[index].AutoHeight = false;
                    settings.Drawers[index].Height = Math.Max(0.0, height);
                }
            }
        }

        private void RebuildDrawerHeightRows(AssemblyDrawerSettings settings)
        {
            DrawerHeightsPanel.Children.Clear();
            for (int i = 0; i < settings.Drawers.Count; i++)
            {
                AssemblyDrawerDraft drawer = settings.Drawers[i];
                var row = new Grid
                {
                    Tag = i,
                    Margin = new Thickness(0, 0, 0, 6)
                };
                row.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
                row.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1.0, GridUnitType.Star) });
                row.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });

                var label = new TextBlock
                {
                    Text = (i + 1).ToString(CultureInfo.InvariantCulture),
                    Foreground = new SolidColorBrush(Color.FromRgb(148, 163, 184)),
                    VerticalAlignment = VerticalAlignment.Center,
                    Margin = new Thickness(0, 0, 8, 0)
                };
                Grid.SetColumn(label, 0);
                row.Children.Add(label);

                var input = new TextBox
                {
                    Text = drawer.AutoHeight ? "auto" : drawer.Height.ToString("0.###", CultureInfo.InvariantCulture),
                    Background = new SolidColorBrush(Color.FromRgb(17, 24, 39)),
                    Foreground = new SolidColorBrush(Color.FromRgb(248, 250, 252)),
                    BorderBrush = new SolidColorBrush(Color.FromRgb(51, 65, 85)),
                    BorderThickness = new Thickness(1),
                    Padding = new Thickness(8, 6, 8, 6),
                    Margin = new Thickness(0, 0, 8, 0)
                };
                input.TextChanged += OnDrawerSettingsChanged;
                Grid.SetColumn(input, 1);
                row.Children.Add(input);

                var autoButton = new Button
                {
                    Content = "auto",
                    Tag = i,
                    Padding = new Thickness(8, 5, 8, 5)
                };
                autoButton.Click += OnDrawerHeightAutoClick;
                Grid.SetColumn(autoButton, 2);
                row.Children.Add(autoButton);

                DrawerHeightsPanel.Children.Add(row);
            }
        }

        private static void ExecuteAcadCommand(string command)
        {
            var doc = AcadApp.DocumentManager.MdiActiveDocument;
            if (doc == null)
            {
                return;
            }

            doc.SendStringToExecute(command, true, false, false);
        }

        private void LoadMaterialChoices()
        {
            materialDatabase = MaterialDatabaseStore.Load();
            string active = MaterialComboBox != null ? MaterialComboBox.Text : null;
            string front = DrawerFrontMaterialComboBox != null ? DrawerFrontMaterialComboBox.Text : null;
            string body = DrawerBodyMaterialComboBox != null ? DrawerBodyMaterialComboBox.Text : null;
            string bottom = DrawerBottomMaterialComboBox != null ? DrawerBottomMaterialComboBox.Text : null;

            object[] boardNames = materialDatabase.BoardMaterials
                .Select(x => (object)MaterialDatabaseStore.DisplayName(x))
                .Where(x => !string.IsNullOrWhiteSpace((string)x))
                .ToArray();
            MaterialComboBox.ItemsSource = boardNames;
            DrawerFrontMaterialComboBox.ItemsSource = boardNames;
            DrawerBodyMaterialComboBox.ItemsSource = boardNames;
            DrawerBottomMaterialComboBox.ItemsSource = boardNames;

            if (!string.IsNullOrWhiteSpace(active)) MaterialComboBox.Text = active;
            if (!string.IsNullOrWhiteSpace(front)) DrawerFrontMaterialComboBox.Text = front;
            if (!string.IsNullOrWhiteSpace(body)) DrawerBodyMaterialComboBox.Text = body;
            if (!string.IsNullOrWhiteSpace(bottom)) DrawerBottomMaterialComboBox.Text = bottom;
        }

        private void ApplyActiveMaterialFromCombo(bool updateThickness)
        {
            string material = MaterialComboBox.Text;
            controller.SetMaterial(material);
            if (!updateThickness)
            {
                return;
            }

            double thickness = ResolveMaterialThickness(material);
            if (thickness > 0.0)
            {
                ThicknessTextBox.Text = thickness.ToString("0.###", CultureInfo.InvariantCulture);
                controller.SetThickness(thickness);
            }
        }

        private double ResolveMaterialThickness(string displayName)
        {
            if (materialDatabase == null)
            {
                materialDatabase = MaterialDatabaseStore.Load();
            }

            BoardMaterialData material = materialDatabase.BoardMaterials.FirstOrDefault(x =>
                string.Equals(MaterialDatabaseStore.DisplayName(x), displayName, StringComparison.CurrentCultureIgnoreCase)
                || string.Equals(x.Code, displayName, StringComparison.CurrentCultureIgnoreCase)
                || string.Equals(x.Name, displayName, StringComparison.CurrentCultureIgnoreCase));
            MaterialFormatData format = material != null ? material.Formats.FirstOrDefault(x => x.Thickness > 0.0) : null;
            return format != null ? format.Thickness : 0.0;
        }

        private static string Safe(string value)
        {
            return string.IsNullOrWhiteSpace(value) ? "-" : value.Trim();
        }
    }
}
