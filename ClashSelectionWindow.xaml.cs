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
                if (test is not ClashTest clashTest)
                    continue;

                foreach (var child in clashTest.Children)
                {
                    if (child is ClashResult clashResult)
                    {
                        items.Add(new ClashItem
                        {
                            Name = clashResult.DisplayName,
                            Status = clashResult.Status.ToString()
                        });
                    }
                    else if (child is ClashResultGroup group)
                    {
                        foreach (var result in GetResultsInGroup(group))
                        {
                            items.Add(new ClashItem
                            {
                                Name = result.DisplayName,
                                Status = result.Status.ToString()
                            });
                        }
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

