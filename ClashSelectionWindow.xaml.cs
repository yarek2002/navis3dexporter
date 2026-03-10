using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using Autodesk.Navisworks.Api;
using Autodesk.Navisworks.Api.Clash;

namespace Navis3dExporter
{
    public partial class ClashSelectionWindow : Window
    {
        private readonly Document _document;

        public ClashSelectionWindow(Document document)
        {
            _document = document ?? throw new ArgumentNullException(nameof(document));
            InitializeComponent();
            LoadClashes();
        }

        private class ClashItem
        {
            public string Name { get; set; }
            public string Status { get; set; }
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

            var items = new List<ClashItem>();

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
                        items.Add(new ClashItem
                        {
                            Name = clashResult.DisplayName,
                            Status = clashResult.Status.ToString()
                        });
                    }
                    // Группы коллизий – отображаем сами группы как в Clash Detective
                    else if (child is ClashResultGroup group)
                    {
                        items.Add(new ClashItem
                        {
                            Name = group.DisplayName,
                            Status = group.Status.ToString()
                        });
                    }
                }
            }

            ClashesGrid.ItemsSource = items;
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

