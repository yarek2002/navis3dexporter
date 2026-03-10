using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using Autodesk.Navisworks.Api;
using Autodesk.Navisworks.Api.Clash;

namespace Navis3dExporter
{
    public partial class ClashSelectionWindow : Window
    {
        private readonly Document _document;
        private readonly string _outputFolder;
        private readonly List<ClashItem> _items = new List<ClashItem>();

        public IReadOnlyList<ClashSelection> SelectedClashes =>
            _items
                .Where(x => x.IsSelected && x.Selection != null)
                .Select(x => x.Selection)
                .ToList();

        public ClashSelectionWindow(Document document, string outputFolder)
        {
            _document = document ?? throw new ArgumentNullException(nameof(document));
            _outputFolder = outputFolder ?? throw new ArgumentNullException(nameof(outputFolder));
            InitializeComponent();
            LoadClashes();
        }

        public class ClashSelection
        {
            public string TestDisplayName { get; set; }
            public string ItemDisplayName { get; set; }
            public ClashSelectionKind Kind { get; set; }
        }

        public enum ClashSelectionKind
        {
            Group = 1,
            Single = 2
        }

        private class ClashItem : INotifyPropertyChanged
        {
            public string Name { get; set; }
            public string Status { get; set; }

            private bool _isSelected;
            public bool IsSelected
            {
                get => _isSelected;
                set
                {
                    if (_isSelected == value) return;
                    _isSelected = value;
                    PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(IsSelected)));
                }
            }

            public ClashSelection Selection { get; set; }

            public event PropertyChangedEventHandler PropertyChanged;
        }

        private void LoadClashes()
        {
            var clashDoc = _document.GetClash();
            if (clashDoc == null || clashDoc.TestsData.Tests.Count == 0)
            {
                MessageBox.Show(
                    this,
                    "В документе нет тестов коллизий (Clash Detective).",
                    "GLB Exporter",
                    MessageBoxButton.OK,
                    MessageBoxImage.Information);
                return;
            }

            _items.Clear();

            foreach (var test in clashDoc.TestsData.Tests)
            {
                var clashTest = test as ClashTest;
                if (clashTest == null)
                    continue;

                foreach (var child in clashTest.Children)
                {
                    // Негруппированные результаты коллизий
                    if (child is ClashResult clashResult)
                    {
                        _items.Add(new ClashItem
                        {
                            Name = clashResult.DisplayName,
                            Status = clashResult.Status.ToString(),
                            Selection = new ClashSelection
                            {
                                TestDisplayName = clashTest.DisplayName,
                                ItemDisplayName = clashResult.DisplayName,
                                Kind = ClashSelectionKind.Single
                            }
                        });
                    }
                    // Группы коллизий – отображаем сами группы как в Clash Detective
                    else if (child is ClashResultGroup group)
                    {
                        _items.Add(new ClashItem
                        {
                            Name = group.DisplayName,
                            Status = group.Status.ToString(),
                            Selection = new ClashSelection
                            {
                                TestDisplayName = clashTest.DisplayName,
                                ItemDisplayName = group.DisplayName,
                                Kind = ClashSelectionKind.Group
                            }
                        });
                    }
                }
            }

            ClashesGrid.ItemsSource = _items;
        }

        private int? _lastClickedIndex;

        private void ClashesGrid_PreviewMouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            var row = FindParent<DataGridRow>(e.OriginalSource as DependencyObject);
            if (row == null)
                return;

            int index = ClashesGrid.ItemContainerGenerator.IndexFromContainer(row);
            if (index < 0 || index >= ClashesGrid.Items.Count)
                return;

            // Выделение диапазона по Shift
            if ((Keyboard.Modifiers & ModifierKeys.Shift) == ModifierKeys.Shift && _lastClickedIndex.HasValue)
            {
                int start = Math.Min(_lastClickedIndex.Value, index);
                int end = Math.Max(_lastClickedIndex.Value, index);

                for (int i = start; i <= end; i++)
                {
                    if (ClashesGrid.Items[i] is ClashItem itemInRange)
                        itemInRange.IsSelected = true;
                }

                e.Handled = true;
            }
            else
            {
                if (row.Item is ClashItem item)
                    item.IsSelected = !item.IsSelected;

                _lastClickedIndex = index;
            }
        }

        private void ExportSelectedButton_OnClick(object sender, RoutedEventArgs e)
        {
            var selected = SelectedClashes;
            if (selected == null || selected.Count == 0)
            {
                MessageBox.Show(
                    this,
                    "Не выбрано ни одной коллизии или группы для экспорта.",
                    "GLB Exporter",
                    MessageBoxButton.OK,
                    MessageBoxImage.Information);
                return;
            }

            try
            {
                var exporter = new GlbClashExporter(_document);
                exporter.ExportSelectedClashes(_outputFolder, selected);

                MessageBox.Show(
                    this,
                    "Экспорт выбранных коллизий в GLB завершён.",
                    "GLB Exporter",
                    MessageBoxButton.OK,
                    MessageBoxImage.Information);

                Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    this,
                    "Ошибка при экспорте выбранных коллизий:\n" + ex.Message,
                    "GLB Exporter",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error);
            }
        }

        private static T FindParent<T>(DependencyObject child) where T : DependencyObject
        {
            while (child != null && !(child is T))
            {
                child = VisualTreeHelper.GetParent(child);
            }

            return child as T;
        }

        private static IEnumerable<ClashResult> GetResultsInGroup(ClashResultGroup group)
        {
            var stack = new Stack<SavedItem>();
            foreach (var child in group.Children)
                stack.Push(child);

            while (stack.Count > 0)
            {
                var item = stack.Pop();
                if (item is ClashResult clashResult)
                {
                    yield return clashResult;
                }
                else if (item is ClashResultGroup subgroup)
                {
                    foreach (var child in subgroup.Children)
                        stack.Push(child);
                }
            }
        }
    }
}

