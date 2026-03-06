using System;
using System.IO;
using System.Windows;
using System.Windows.Forms;

namespace Navis3dExporter
{
    public partial class ExportWindow : Window
    {
        public string SelectedFolder { get; private set; }
        public bool ExportWholeModel => WholeModelModeRadio.IsChecked == true;

        public ExportWindow()
        {
            InitializeComponent();
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
    }
}

