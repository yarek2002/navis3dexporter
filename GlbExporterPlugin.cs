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
                Document doc = NavisApp.ActiveDocument;
                if (doc == null || doc.Models.Count == 0)
                {
                    System.Windows.MessageBox.Show(
                        "Нет загруженной модели для экспорта.",
                        "GLB Exporter");
                    return 0;
                }

                var window = new ExportWindow();
                // Показываем окно модально
                var result = window.ShowDialog();
                if (result != true || string.IsNullOrWhiteSpace(window.SelectedFolder))
                {
                    return 0;
                }

                var exporter = new GlbClashExporter(doc);
                exporter.ExportAllClashes(window.SelectedFolder);

                System.Windows.MessageBox.Show(
                    "Экспорт коллизий в GLB завершён.",
                    "GLB Exporter");

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

