using System;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using AcApplication = Autodesk.AutoCAD.ApplicationServices.Application;

namespace AutoCAD_BoardSorter.Ui
{
    internal sealed class AssemblyPaletteControl : UserControl
    {
        public AssemblyPaletteControl()
        {
            var root = new DockPanel
            {
                LastChildFill = true,
                Background = new SolidColorBrush(Color.FromRgb(42, 48, 56))
            };

            var scroll = new ScrollViewer
            {
                VerticalScrollBarVisibility = ScrollBarVisibility.Auto
            };
            root.Children.Add(scroll);

            var stack = new StackPanel
            {
                Margin = new Thickness(10, 10, 10, 12)
            };
            scroll.Content = stack;

            stack.Children.Add(new TextBlock
            {
                Text = "Инструменты сборок",
                Foreground = Brushes.White,
                FontSize = 16,
                FontWeight = FontWeights.SemiBold,
                Margin = new Thickness(0, 0, 0, 4)
            });

            stack.Children.Add(new TextBlock
            {
                Text = "Быстрый доступ к контейнерам, изоляции и запуску конструктора сборок.",
                Foreground = Brushes.Gainsboro,
                TextWrapping = TextWrapping.Wrap,
                Margin = new Thickness(0, 0, 0, 12)
            });

            stack.Children.Add(CreateSectionHeader("Контейнеры"));
            stack.Children.Add(CreateCommandButton("Сделать контейнер", "BDMAKECONTAINER"));
            stack.Children.Add(CreateCommandButton("Скрыть контейнеры", "BDHIDECONTAINERS"));
            stack.Children.Add(CreateCommandButton("Показать контейнеры", "BDSHOWCONTAINERS"));
            stack.Children.Add(CreateCommandButton("Изолировать детали сборки", "BDISOLATEASSEMBLY"));
            stack.Children.Add(CreateCommandButton("Показать все", "BDSHOWALL"));

            stack.Children.Add(CreateSectionHeader("Черчение"));
            stack.Children.Add(CreateCommandButton("База материалов", "BDMATERIALS"));
            stack.Children.Add(CreateCommandButton("Открыть конструктор сборок", "BDASSEMBLYEDITOR"));

            Content = root;
        }

        private static TextBlock CreateSectionHeader(string text)
        {
            return new TextBlock
            {
                Text = text,
                Foreground = Brushes.White,
                FontWeight = FontWeights.Bold,
                Margin = new Thickness(0, 10, 0, 8)
            };
        }

        private static Button CreateCommandButton(string text, string command)
        {
            var button = new Button
            {
                Content = text,
                HorizontalContentAlignment = HorizontalAlignment.Left,
                Margin = new Thickness(0, 0, 0, 8),
                Padding = new Thickness(10, 8, 10, 8),
                MinHeight = 34
            };

            button.Click += delegate
            {
                var doc = AcApplication.DocumentManager.MdiActiveDocument;
                if (doc == null)
                {
                    return;
                }

                doc.SendStringToExecute(command + " ", true, false, false);
            };

            return button;
        }
    }
}
