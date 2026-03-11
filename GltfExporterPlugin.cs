using System;
using Autodesk.Navisworks.Api;
using Autodesk.Navisworks.Api.Plugins;
using System.Windows;

using NavisApp = Autodesk.Navisworks.Api.Application;

namespace Navis3dExporter
{
    [Plugin("Navis3dExporter.GltfExporter", "NV3D",
        DisplayName = "glTF Exporter",
        ToolTip = "Экспорт текущей модели Navisworks в формат glTF")]
    public class GltfExporterPlugin : AddInPlugin
    {
        public override int Execute(params string[] parameters)
        {
            try
            {
                AssemblyResolver.EnsureRegistered();

                Document doc = NavisApp.ActiveDocument;
                if (doc == null || doc.Models.Count == 0)
                {
                    System.Windows.MessageBox.Show(
                        "Нет загруженной модели для экспорта.",
                        "glTF Exporter");
                    return 0;
                }

                var window = new ExportWindow(doc);
                window.Title = "Экспорт в glTF";
                // Показываем окно модально
                var result = window.ShowDialog();
                if (result != true || string.IsNullOrWhiteSpace(window.SelectedFolder))
                {
                    return 0;
                }

                var exporter = new GlbClashExporter(doc);

                if (window.ExportWholeModel)
                {
                    // Экспорт всей модели в один glTF-файл.
                    var filePath = System.IO.Path.Combine(window.SelectedFolder, "Model.gltf");
                    exporter.ExportWholeModelAsGltf(filePath);

                    System.Windows.MessageBox.Show(
                        "Экспорт всей модели в glTF завершён.",
                        "glTF Exporter");
                }
                else if (window.ExportSelection)
                {
                    // Экспорт текущего выделения в один glTF-файл.
                    var filePath = System.IO.Path.Combine(window.SelectedFolder, "SelectedObjects.gltf");
                    exporter.ExportCurrentSelectionAsGltf(filePath);

                    System.Windows.MessageBox.Show(
                        "Экспорт выделенных объектов в glTF завершён.",
                        "glTF Exporter");
                }
                else
                {
                    // Экспорт всех коллизий из Clash Detective.
                    // Для каждого теста создаётся папка, в ней по glTF на группу/отдельный результат.
                    var clashDoc = doc.GetClash();
                    if (clashDoc == null || clashDoc.TestsData.Tests.Count == 0)
                    {
                        System.Windows.MessageBox.Show(
                            "В документе нет тестов коллизий (Clash Detective).",
                            "glTF Exporter");
                        return 0;
                    }

                    System.Windows.MessageBox.Show(
                        "Экспорт коллизий в glTF пока не реализован. Используйте GLB Exporter для экспорта коллизий.",
                        "glTF Exporter");
                }

                return 0;
            }
            catch (Exception ex)
            {
                System.Windows.MessageBox.Show(
                    "Ошибка в плагине glTF Exporter:\n" + ex.Message,
                    "glTF Exporter");
                return -1;
            }
        }
    }
}

