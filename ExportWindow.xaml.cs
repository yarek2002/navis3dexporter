using System;
using System.Collections.Generic;
using System.IO;
using System.Windows;
using System.Windows.Forms;
using Autodesk.Navisworks.Api;

namespace Navis3dExporter
{
    public partial class ExportWindow : Window
    {
        public string SelectedFolder { get; private set; }
        public bool ExportWholeModel => WholeModelModeRadio.IsChecked == true;
        public bool ExportSelection => SelectionModeRadio.IsChecked == true;
        private readonly Document _document;

        public ExportWindow(Document document)
        {
            _document = document ?? throw new ArgumentNullException(nameof(document));
            InitializeComponent();
            this.Width = 1200;
            this.Height = 200;

        }

        private void BrowseButton_OnClick(object sender, RoutedEventArgs e)
        {
            using (var dialog = new FolderBrowserDialog())
            {
                dialog.Description = "Выберите папку для сохранения GLB файлов коллизий";
                dialog.ShowNewFolderButton = true;

                if (!string.IsNullOrWhiteSpace(FolderTextBox.Text) &&
                    Directory.Exists(FolderTextBox.Text))
                {
                    dialog.SelectedPath = FolderTextBox.Text;
                }

                if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK &&
                    !string.IsNullOrWhiteSpace(dialog.SelectedPath))
                {
                    FolderTextBox.Text = dialog.SelectedPath;
                }
            }
        }

        private void RunButton_OnClick(object sender, RoutedEventArgs e)
        {
            var path = FolderTextBox.Text?.Trim();
            if (string.IsNullOrWhiteSpace(path))
            {
                System.Windows.MessageBox.Show(
                    this,
                    "Укажите папку для сохранения результатов.",
                    "GLB Exporter",
                    MessageBoxButton.OK,
                    MessageBoxImage.Warning);
                return;
            }

            try
            {
                Directory.CreateDirectory(path);
            }
            catch (Exception ex)
            {
                System.Windows.MessageBox.Show(
                    this,
                    "Не удалось создать/открыть папку:\n" + ex.Message,
                    "GLB Exporter",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error);
                return;
            }

            SelectedFolder = path;
            DialogResult = true;
            Close();
        }

        private void SelectClashesButton_OnClick(object sender, RoutedEventArgs e)
        {
            var path = FolderTextBox.Text?.Trim();
            if (string.IsNullOrWhiteSpace(path))
            {
                System.Windows.MessageBox.Show(
                    this,
                    "Укажите папку для сохранения результатов перед выбором коллизий.",
                    "GLB Exporter",
                    MessageBoxButton.OK,
                    MessageBoxImage.Warning);
                return;
            }

            try
            {
                Directory.CreateDirectory(path);
            }
            catch (Exception ex)
            {
                System.Windows.MessageBox.Show(
                    this,
                    "Не удалось создать/открыть папку:\n" + ex.Message,
                    "GLB Exporter",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error);
                return;
            }

            try
            {
                var wnd = new ClashSelectionWindow(_document, path)
                {
                    Owner = this
                };
                wnd.ShowDialog();
            }
            catch (Exception ex)
            {
                System.Windows.MessageBox.Show(
                    this,
                    "Не удалось загрузить список коллизий:\n" + ex.Message,
                    "GLB Exporter",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error);
            }
        }

    }
}

