using System;
using Autodesk.Navisworks.Api;
using Autodesk.Navisworks.Api.Plugins;
using System.Windows;

using NavisApp = Autodesk.Navisworks.Api.Application;

namespace Navis3dExporter
{
    [Plugin("Navis3dExporter.GlbExporter", "NV3D",
        DisplayName = "GLB Exporter",
        ToolTip = "Экспорт текущей модели Navisworks в формат GLB")]
    public class GlbExporterPlugin : AddInPlugin
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
                        "GLB Exporter");
                    return 0;
                }

                var window = new ExportWindow(doc);
                // Показываем окно модально
                var result = window.ShowDialog();
                if (result != true || string.IsNullOrWhiteSpace(window.SelectedFolder))
                {
                    return 0;
                }

                var exporter = new GlbClashExporter(doc);

                if (window.ExportWholeModel)
                {
                    // Экспорт всей модели в один GLB-файл.
                    var filePath = System.IO.Path.Combine(window.SelectedFolder, "Model.glb");
                    exporter.ExportWholeModel(filePath);

                    System.Windows.MessageBox.Show(
                        "Экспорт всей модели в GLB завершён.",
                        "GLB Exporter");
                }
                else
                {
                    // Экспорт всех коллизий из Clash Detective:
                    // для каждого теста создаётся папка, в ней по GLB на группу/отдельный результат.
                    exporter.ExportAllClashes(window.SelectedFolder);

                    System.Windows.MessageBox.Show(
                        "Экспорт всех групп коллизий в GLB завершён.",
                        "GLB Exporter");
                }

                return 0;
            }
            catch (Exception ex)
            {
                System.Windows.MessageBox.Show(
                    "Ошибка в плагине GLB Exporter:\n" + ex.Message,
                    "GLB Exporter");
                return -1;
            }
        }
    }
}

