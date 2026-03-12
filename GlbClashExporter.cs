using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Numerics;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;
using Autodesk.Navisworks.Api;
using Autodesk.Navisworks.Api.Clash;
using SharpGLTF.Geometry;
using SharpGLTF.Geometry.VertexTypes;
using SharpGLTF.Materials;
using SharpGLTF.Scenes;
using SharpGLTF.Schema2;
using ComBridge = Autodesk.Navisworks.Api.ComApi.ComApiBridge;
using COMApi = Autodesk.Navisworks.Api.Interop.ComApi;

namespace Navis3dExporter
{
    internal class GlbClashExporter
    {
        private readonly Document _document;

        // Navisworks использует Z-up, glTF — Y-up.
        // Поворот на -90° вокруг оси X переводит систему координат Navisworks в систему glTF.
        private static readonly Matrix4x4 NavisToGltfTransform =
            Matrix4x4.CreateRotationX(-(float)(Math.PI / 2));

        public GlbClashExporter(Document document)
        {
            _document = document ?? throw new ArgumentNullException(nameof(document));
        }

        public void ExportWholeModel(string outputFilePath)
        {
            if (string.IsNullOrWhiteSpace(outputFilePath))
                throw new ArgumentException("Output file path is not specified.", nameof(outputFilePath));

            var roots = _document.Models?.RootItems;
            if (roots == null)
                throw new InvalidOperationException("В документе нет корневых элементов модели для экспорта.");

            var scene = new SceneBuilder();

            int index = 0;
            foreach (ModelItem root in roots)
            {
                index++;

                var mesh = BuildMeshFromModelItem(
                    root,
                    $"Root_{index}",
                    new Vector4(0.8f, 0.8f, 0.8f, 1f));

                if (mesh != null)
                {
                    scene.AddRigidMesh(mesh, NavisToGltfTransform);
                }
            }

            if (index == 0)
                throw new InvalidOperationException("В документе нет корневых элементов модели для экспорта.");

            var model = scene.ToGltf2();
            model.SaveGLB(outputFilePath);
        }

        public void ExportWholeModelAsGltf(string outputFilePath)
        {
            if (string.IsNullOrWhiteSpace(outputFilePath))
                throw new ArgumentException("Output file path is not specified.", nameof(outputFilePath));

            var roots = _document.Models?.RootItems;
            if (roots == null)
                throw new InvalidOperationException("В документе нет корневых элементов модели для экспорта.");

            var scene = new SceneBuilder();

            int index = 0;
            foreach (ModelItem root in roots)
            {
                index++;

                var mesh = BuildMeshFromModelItem(
                    root,
                    $"Root_{index}",
                    new Vector4(0.8f, 0.8f, 0.8f, 1f));

                if (mesh != null)
                {
                    scene.AddRigidMesh(mesh, NavisToGltfTransform);
                }
            }

            if (index == 0)
                throw new InvalidOperationException("В документе нет корневых элементов модели для экспорта.");

            var model = scene.ToGltf2();
            var settings = new WriteSettings { MergeBuffers = false };
            model.SaveGLTF(outputFilePath, settings);
        }

        public void ExportCurrentSelection(string outputFilePath)
        {
            if (string.IsNullOrWhiteSpace(outputFilePath))
                throw new ArgumentException("Output file path is not specified.", nameof(outputFilePath));

            var selection = _document.CurrentSelection;
            if (selection == null || selection.SelectedItems == null || selection.SelectedItems.Count == 0)
                throw new InvalidOperationException("В текущем выделении нет элементов для экспорта.");

            var scene = new SceneBuilder();

            int total = selection.SelectedItems.Count;

            // Если выделено ровно два элемента – сохраняем старое поведение:
            // первый красный, второй синий.
            if (total == 2)
            {
                int index = 0;
                foreach (ModelItem item in selection.SelectedItems)
                {
                    index++;

                    var color = (index == 1)
                        ? new Vector4(1f, 0f, 0f, 1f)   // красный
                        : new Vector4(0f, 0f, 1f, 1f);  // синий

                    var mesh = BuildMeshFromModelItem(item, $"Sel_{index}", color);
                    if (mesh != null)
                    {
                        scene.AddRigidMesh(mesh, NavisToGltfTransform);
                    }
                }
            }
            else
            {
                // Если элементов больше двух – красим по исходной модели (NWC):
                // все элементы из одной NWC получают один цвет.
                var rootColorMap = BuildRootColorMap();

                int index = 0;
                foreach (ModelItem item in selection.SelectedItems)
                {
                    index++;

                    var color = GetColorForModelItem(item, rootColorMap, 0);

                    var mesh = BuildMeshFromModelItem(item, $"Sel_{index}", color);
                    if (mesh != null)
                    {
                        scene.AddRigidMesh(mesh, NavisToGltfTransform);
                    }
                }
            }

            var model = scene.ToGltf2();
            model.SaveGLB(outputFilePath);
        }

        public void ExportCurrentSelectionAsGltf(string outputFilePath)
        {
            if (string.IsNullOrWhiteSpace(outputFilePath))
                throw new ArgumentException("Output file path is not specified.", nameof(outputFilePath));

            var selection = _document.CurrentSelection;
            if (selection == null || selection.SelectedItems == null || selection.SelectedItems.Count == 0)
                throw new InvalidOperationException("В текущем выделении нет элементов для экспорта.");

            var scene = new SceneBuilder();

            int total = selection.SelectedItems.Count;

            if (total == 2)
            {
                int index = 0;
                foreach (ModelItem item in selection.SelectedItems)
                {
                    index++;

                    var color = (index == 1)
                        ? new Vector4(1f, 0f, 0f, 1f)
                        : new Vector4(0f, 0f, 1f, 1f);

                    var mesh = BuildMeshFromModelItem(item, $"Sel_{index}", color);
                    if (mesh != null)
                    {
                        scene.AddRigidMesh(mesh, NavisToGltfTransform);
                    }
                }
            }
            else
            {
                var rootColorMap = BuildRootColorMap();

                int index = 0;
                foreach (ModelItem item in selection.SelectedItems)
                {
                    index++;

                    var color = GetColorForModelItem(item, rootColorMap, 0);

                    var mesh = BuildMeshFromModelItem(item, $"Sel_{index}", color);
                    if (mesh != null)
                    {
                        scene.AddRigidMesh(mesh, NavisToGltfTransform);
                    }
                }
            }

            var model = scene.ToGltf2();
            var settings = new WriteSettings { MergeBuffers = false };
            model.SaveGLTF(outputFilePath, settings);
        }

        public void ExportAllClashes(string outputFolder)
        {
            if (string.IsNullOrWhiteSpace(outputFolder))
                throw new ArgumentException("Output folder is not specified.", nameof(outputFolder));

            var clashDoc = _document.GetClash();
            if (clashDoc == null || clashDoc.TestsData.Tests.Count == 0)
                throw new InvalidOperationException("В документе нет тестов коллизий (Clash Detective).");

            Directory.CreateDirectory(outputFolder);

            var tests = clashDoc.TestsData.Tests;
            int testIndex = 0;

            foreach (var test in tests)
            {
                var clashTest = test as ClashTest;
                if (clashTest == null)
                    continue;

                testIndex++;

                string testFolderName = BuildSafeNamedSegment(
                    displayName: clashTest.DisplayName,
                    maxSegmentLen: 60);

                string testFolder = Path.Combine(outputFolder, testFolderName);
                Directory.CreateDirectory(testFolder);

                int groupIndex = 0;
                int clashIndex = 0;

                foreach (var child in clashTest.Children)
                {
                    if (child is ClashResultGroup group)
                    {
                        groupIndex++;

                        string groupName = BuildSafeNamedSegment(
                            displayName: ExtractTrailingGuidToken(group.DisplayName),
                            maxSegmentLen: 80);

                        string filePath = Path.Combine(testFolder, groupName + ".glb");
                        ExportGroup(group, filePath);
                    }
                    else if (child is ClashResult clashResult)
                    {
                        // Одиночные результаты без группы
                        clashIndex++;
                        string clashName = BuildSafeNamedSegment(
                            displayName: ExtractTrailingGuidToken(clashResult.DisplayName),
                            maxSegmentLen: 80);

                        string filePath = Path.Combine(testFolder, clashName + ".glb");
                        ExportSingleClash(clashResult, filePath);
                    }
                }
            }
        }

        public void ExportSelectedClashes(
            string outputFolder,
            IReadOnlyList<ClashSelectionWindow.ClashSelection> selected)
        {
            if (string.IsNullOrWhiteSpace(outputFolder))
                throw new ArgumentException("Output folder is not specified.", nameof(outputFolder));
            if (selected == null)
                throw new ArgumentNullException(nameof(selected));
            if (selected.Count == 0)
                throw new InvalidOperationException("Не выбрано ни одной коллизии/группы для экспорта.");

            var clashDoc = _document.GetClash();
            if (clashDoc == null || clashDoc.TestsData.Tests.Count == 0)
                throw new InvalidOperationException("В документе нет тестов коллизий (Clash Detective).");

            Directory.CreateDirectory(outputFolder);

            // Сопоставление по DisplayName (как видно в Clash Detective).
            // Для имен файлов используем ExtractTrailingGuidToken + BuildSafeNamedSegment, как и в ExportAllClashes.
            var byTest = selected
                .GroupBy(x => x.TestDisplayName ?? string.Empty)
                .ToDictionary(g => g.Key, g => g.ToList());

            foreach (var test in clashDoc.TestsData.Tests)
            {
                var clashTest = test as ClashTest;
                if (clashTest == null)
                    continue;

                if (!byTest.TryGetValue(clashTest.DisplayName ?? string.Empty, out var pickedInTest) ||
                    pickedInTest.Count == 0)
                    continue;

                string testFolderName = BuildSafeNamedSegment(
                    displayName: clashTest.DisplayName,
                    maxSegmentLen: 60);

                string testFolder = Path.Combine(outputFolder, testFolderName);
                Directory.CreateDirectory(testFolder);

                var pickedGroups = new HashSet<string>(
                    pickedInTest
                        .Where(x => x.Kind == ClashSelectionWindow.ClashSelectionKind.Group)
                        .Select(x => x.ItemDisplayName ?? string.Empty));

                var pickedSingles = new HashSet<string>(
                    pickedInTest
                        .Where(x => x.Kind == ClashSelectionWindow.ClashSelectionKind.Single)
                        .Select(x => x.ItemDisplayName ?? string.Empty));

                foreach (var child in clashTest.Children)
                {
                    if (child is ClashResultGroup group)
                    {
                        if (!pickedGroups.Contains(group.DisplayName ?? string.Empty))
                            continue;

                        string groupName = BuildSafeNamedSegment(
                            displayName: ExtractTrailingGuidToken(group.DisplayName),
                            maxSegmentLen: 80);

                        string filePath = Path.Combine(testFolder, groupName + ".glb");
                        ExportGroup(group, filePath);
                    }
                    else if (child is ClashResult clashResult)
                    {
                        if (!pickedSingles.Contains(clashResult.DisplayName ?? string.Empty))
                            continue;

                        string clashName = BuildSafeNamedSegment(
                            displayName: ExtractTrailingGuidToken(clashResult.DisplayName),
                            maxSegmentLen: 80);

                        string filePath = Path.Combine(testFolder, clashName + ".glb");
                        ExportSingleClash(clashResult, filePath);
                    }
                }
            }
        }

        public void ExportSelectedClashesToSingleFile(
            string outputFilePath,
            IReadOnlyList<ClashSelectionWindow.ClashSelection> selected)
        {
            if (string.IsNullOrWhiteSpace(outputFilePath))
                throw new ArgumentException("Output file path is not specified.", nameof(outputFilePath));
            if (selected == null)
                throw new ArgumentNullException(nameof(selected));
            if (selected.Count == 0)
                throw new InvalidOperationException("Не выбрано ни одной коллизии/группы для экспорта.");

            var clashDoc = _document.GetClash();
            if (clashDoc == null || clashDoc.TestsData.Tests.Count == 0)
                throw new InvalidOperationException("В документе нет тестов коллизий (Clash Detective).");

            var byTest = selected
                .GroupBy(x => x.TestDisplayName ?? string.Empty)
                .ToDictionary(g => g.Key, g => g.ToList());

            var scene = new SceneBuilder();
            int clashIndex = 0;

            // Кэш цветов по корневым моделям (каждая NWC получает свой цвет).
            var rootColorMap = BuildRootColorMap();

            foreach (var test in clashDoc.TestsData.Tests)
            {
                var clashTest = test as ClashTest;
                if (clashTest == null)
                    continue;

                if (!byTest.TryGetValue(clashTest.DisplayName ?? string.Empty, out var pickedInTest) ||
                    pickedInTest.Count == 0)
                    continue;

                var pickedGroups = new HashSet<string>(
                    pickedInTest
                        .Where(x => x.Kind == ClashSelectionWindow.ClashSelectionKind.Group)
                        .Select(x => x.ItemDisplayName ?? string.Empty));

                var pickedSingles = new HashSet<string>(
                    pickedInTest
                        .Where(x => x.Kind == ClashSelectionWindow.ClashSelectionKind.Single)
                        .Select(x => x.ItemDisplayName ?? string.Empty));

                foreach (var child in clashTest.Children)
                {
                    if (child is ClashResultGroup group)
                    {
                        if (!pickedGroups.Contains(group.DisplayName ?? string.Empty))
                            continue;

                        foreach (var clash in GetResultsInGroup(group))
                        {
                            clashIndex++;
                            AddClashToScene(scene, clash, clashIndex, rootColorMap);
                        }
                    }
                    else if (child is ClashResult clashResult)
                    {
                        if (!pickedSingles.Contains(clashResult.DisplayName ?? string.Empty))
                            continue;

                        clashIndex++;
                        AddClashToScene(scene, clashResult, clashIndex, rootColorMap);
                    }
                }
            }

            if (clashIndex == 0)
                throw new InvalidOperationException("Не удалось найти выбранные коллизии для экспорта.");

            var model = scene.ToGltf2();
            model.SaveGLB(outputFilePath);
        }

        private void AddClashToScene(
            SceneBuilder scene,
            ClashResult clash,
            int index,
            Dictionary<ModelItem, Vector4> rootColorMap)
        {
            if (scene == null) throw new ArgumentNullException(nameof(scene));
            if (clash == null) return;

            var item1 = clash.Item1;
            var item2 = clash.Item2;

            if (item1 == null || item2 == null)
                return;

            var color1 = GetColorForModelItem(item1, rootColorMap, 0);
            var color2 = GetColorForModelItem(item2, rootColorMap, 1);

            var mesh1 = BuildMeshFromModelItem(item1, $"Item1_{index}", color1);
            if (mesh1 != null)
                scene.AddRigidMesh(mesh1, NavisToGltfTransform);

            var mesh2 = BuildMeshFromModelItem(item2, $"Item2_{index}", color2);
            if (mesh2 != null)
                scene.AddRigidMesh(mesh2, NavisToGltfTransform);
        }

        private Dictionary<ModelItem, Vector4> BuildRootColorMap()
        {
            var map = new Dictionary<ModelItem, Vector4>();

            var roots = _document.Models?.RootItems;
            if (roots == null)
                return map;

            foreach (ModelItem root in roots)
            {
                var id = GetRootIdentifier(root);
                map[root] = GetColorFromKey(id);
            }

            return map;
        }

        private Vector4 GetColorForModelItem(
            ModelItem item,
            Dictionary<ModelItem, Vector4> rootColorMap,
            int localIndex)
        {
            // Если что-то не так с кэшем, возвращаем базовые цвета.
            if (item == null || rootColorMap == null || rootColorMap.Count == 0)
            {
                return localIndex == 0
                    ? new Vector4(1f, 0f, 0f, 1f)
                    : new Vector4(0f, 0f, 1f, 1f);
            }

            // Ищем корневой элемент (верхний в иерархии).
            var root = item;
            while (root.Parent != null)
                root = root.Parent;

            if (!rootColorMap.TryGetValue(root, out var color))
            {
                var id = GetRootIdentifier(root);
                color = GetColorFromKey(id);
                rootColorMap[root] = color;
            }

            return color;
        }

        private static string GetRootIdentifier(ModelItem root)
        {
            if (root == null) return string.Empty;

            // Используем DisplayName как основной идентификатор модели.
            var name = root.DisplayName;
            if (string.IsNullOrWhiteSpace(name))
            {
                name = root.ClassName;
            }

            return name ?? string.Empty;
        }

        private static Vector4 GetColorFromKey(string key)
        {
            if (string.IsNullOrEmpty(key))
                return new Vector4(0.6f, 0.6f, 0.6f, 1f);

            unchecked
            {
                int hash = key.GetHashCode();
                if (hash < 0) hash = -hash;

                // Преобразуем хеш в оттенок [0..360).
                float h = (hash % 360) / 360f;
                const float s = 0.65f;
                const float v = 0.95f;

                HsvToRgb(h, s, v, out float r, out float g, out float b);
                return new Vector4(r, g, b, 1f);
            }
        }

        private static void HsvToRgb(float h, float s, float v, out float r, out float g, out float b)
        {
            if (s <= 0f)
            {
                r = g = b = v;
                return;
            }

            h = (h - (float)Math.Floor(h)) * 6f;
            int i = (int)Math.Floor(h);
            float f = h - i;
            float p = v * (1f - s);
            float q = v * (1f - s * f);
            float t = v * (1f - s * (1f - f));

            switch (i)
            {
                case 0: r = v; g = t; b = p; break;
                case 1: r = q; g = v; b = p; break;
                case 2: r = p; g = v; b = t; break;
                case 3: r = p; g = q; b = v; break;
                case 4: r = t; g = p; b = v; break;
                default: r = v; g = p; b = q; break;
            }
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

        private void ExportSingleClash(ClashResult clash, string filePath)
        {
            var item1 = clash.Item1;
            var item2 = clash.Item2;

            if (item1 == null || item2 == null)
                return;

            var scene = new SceneBuilder();

            // Первый элемент – красным
            var mesh1 = BuildMeshFromModelItem(item1, "Item1",
                new Vector4(1f, 0f, 0f, 1f));
            if (mesh1 != null)
                scene.AddRigidMesh(mesh1, NavisToGltfTransform);

            // Второй элемент – синим
            var mesh2 = BuildMeshFromModelItem(item2, "Item2",
                new Vector4(0f, 0f, 1f, 1f));
            if (mesh2 != null)
                scene.AddRigidMesh(mesh2, NavisToGltfTransform);

            var model = scene.ToGltf2();
            model.SaveGLB(filePath);
        }

        private void ExportGroup(ClashResultGroup group, string filePath)
        {
            var results = GetResultsInGroup(group).ToList();
            if (results.Count == 0)
                return;

            var scene = new SceneBuilder();
            int index = 0;

            foreach (var clash in results)
            {
                index++;

                var item1 = clash.Item1;
                var item2 = clash.Item2;

                if (item1 == null || item2 == null)
                    continue;

                var mesh1 = BuildMeshFromModelItem(item1, $"Item1_{index}",
                    new Vector4(1f, 0f, 0f, 1f));
                if (mesh1 != null)
                    scene.AddRigidMesh(mesh1, NavisToGltfTransform);

                var mesh2 = BuildMeshFromModelItem(item2, $"Item2_{index}",
                    new Vector4(0f, 0f, 1f, 1f));
                if (mesh2 != null)
                    scene.AddRigidMesh(mesh2, NavisToGltfTransform);
            }

            var model = scene.ToGltf2();
            model.SaveGLB(filePath);
        }

        private static string ExtractTrailingGuidToken(string displayName)
        {
            if (string.IsNullOrWhiteSpace(displayName)) return displayName;

            // Пользовательский формат: значимая часть (GUID) в самом конце, разделитель '|'
            // Пример: "Some text | more info | 3f2504e0-4f89-11d3-9a0c-0305e82c3301"
            var tail = displayName;
            var pipeIndex = displayName.LastIndexOf('|');
            if (pipeIndex >= 0 && pipeIndex < displayName.Length - 1)
            {
                tail = displayName.Substring(pipeIndex + 1).Trim();
            }

            // Иногда GUID дополнительно отделяют '_' (например "...|...|name_<guid>")
            var underscoreIndex = tail.LastIndexOf('_');
            if (underscoreIndex >= 0 && underscoreIndex < tail.Length - 1)
            {
                var afterUnderscore = tail.Substring(underscoreIndex + 1).Trim();
                if (LooksLikeGuid(afterUnderscore)) return afterUnderscore;
            }

            // Если хвост — GUID, используем его; иначе оставляем исходное имя,
            // чтобы не терять информацию при нестандартных форматах.
            return LooksLikeGuid(tail) ? tail : displayName;
        }

        private static bool LooksLikeGuid(string value)
        {
            if (string.IsNullOrWhiteSpace(value)) return false;
            value = value.Trim().Trim('{', '}');

            if (Guid.TryParse(value, out _)) return true;

            return Regex.IsMatch(
                value,
                @"^[0-9a-fA-F]{8}(-[0-9a-fA-F]{4}){3}-[0-9a-fA-F]{12}$");
        }

        private MeshBuilder<VertexPositionNormal, VertexEmpty, VertexEmpty> BuildMeshFromModelItem(
            ModelItem modelItem,
            string meshName,
            Vector4 color)
        {
           var triangles = ExtractTriangles(modelItem);
            if (triangles.Count == 0)
                return null;

            var mesh = new MeshBuilder<VertexPositionNormal, VertexEmpty, VertexEmpty>(meshName);

            var material = new MaterialBuilder()
                .WithDoubleSide(true)
                .WithMetallicRoughnessShader()
                .WithBaseColor(color);

            var prim = mesh.UsePrimitive(material);

            foreach (var tri in triangles)
            {
                var v0 = new VertexPositionNormal(tri.V0, tri.N0);
                var v1 = new VertexPositionNormal(tri.V1, tri.N1);
                var v2 = new VertexPositionNormal(tri.V2, tri.N2);

                prim.AddTriangle(v0, v1, v2);
            }

            return mesh;
        }

        private struct TriangleData
        {
            public Vector3 V0;
            public Vector3 V1;
            public Vector3 V2;
            public Vector3 N0;
            public Vector3 N1;
            public Vector3 N2;
        }

        private List<TriangleData> ExtractTriangles(ModelItem modelItem)
        {
            var triangles = new List<TriangleData>();

            // Копируем проверенный паттерн из примера:
            // 1) строим COM-выборку только из данного ModelItem
            // 2) обходим Paths
            // 3) для каждого path берём только те фрагменты, которые реально относятся к нему
            // 4) для каждого фрагмента вызываем GenerateSimplePrimitives с корректной матрицей LCS->WCS

            var oSel = (COMApi.InwOpSelection)ComBridge.ToInwOpSelection(
                new ModelItemCollection { modelItem });

            var callback = new TriangleCollector(triangles);

            foreach (COMApi.InwOaPath3 path in oSel.Paths())
            {
                var path1 = path;

                var fragments = path.Fragments()
                    .Cast<COMApi.InwOaFragment3>()
                    .Where(f => IsFragmentOnPath(path1, f))
                    .ToList();

                foreach (var fragment in fragments)
                {
                    callback.CurrentTransform = GetLocalToWorldTransformMatrix(fragment);

                    fragment.GenerateSimplePrimitives(
                        COMApi.nwEVertexProperty.eNORMAL,
                        callback);
                }
            }

            return triangles;
        }

        private class TriangleCollector : COMApi.InwSimplePrimitivesCB
        {
            private readonly List<TriangleData> _triangles;

            public double[] CurrentTransform { get; set; } = null;

            public TriangleCollector(List<TriangleData> triangles)
            {
                _triangles = triangles;
            }

            public void Line(COMApi.InwSimpleVertex v1, COMApi.InwSimpleVertex v2)
            {
                // Не используется
            }

            public void Point(COMApi.InwSimpleVertex v1)
            {
                // Не используется
            }

            public void SnapPoint(COMApi.InwSimpleVertex v1)
            {
                // Не используется
            }

            public void Triangle(
                COMApi.InwSimpleVertex v1,
                COMApi.InwSimpleVertex v2,
                COMApi.InwSimpleVertex v3)
            {
                var p0 = GetPoint(v1, CurrentTransform);
                var p1 = GetPoint(v2, CurrentTransform);
                var p2 = GetPoint(v3, CurrentTransform);

                // Нормали берём из Navisworks (eNORMAL), чтобы:
                // - сохранять сглаживание
                // - повысить шанс переиспользования вершин (welding) в glTF
                var n0 = GetNormal(v1, CurrentTransform);
                var n1 = GetNormal(v2, CurrentTransform);
                var n2 = GetNormal(v3, CurrentTransform);

                // Если нормали не удалось получить, fallback на нормаль грани
                if (n0 == Vector3.Zero || n1 == Vector3.Zero || n2 == Vector3.Zero)
                {
                    var face = Vector3.Normalize(Vector3.Cross(p1 - p0, p2 - p0));
                    if (n0 == Vector3.Zero) n0 = face;
                    if (n1 == Vector3.Zero) n1 = face;
                    if (n2 == Vector3.Zero) n2 = face;
                }

                _triangles.Add(new TriangleData
                {
                    V0 = p0,
                    V1 = p1,
                    V2 = p2,
                    N0 = n0,
                    N1 = n1,
                    N2 = n2
                });
            }

            private static Vector3 GetPoint(COMApi.InwSimpleVertex vertex, double[] matrix)
            {
                // Логика эквивалентна примеру из primer:
                // берём coord как float[3] и умножаем на 4x4-матрицу LCS->WCS.
                var arr = (Array)vertex.coord;
                var v = arr.Cast<float>().ToArray();

                if (matrix == null || matrix.Length != 16)
                {
                    return new Vector3(v[0], v[1], v[2]);
                }

                var x = v[0] * matrix[0] + v[1] * matrix[4] + v[2] * matrix[8] + matrix[12];
                var y = v[0] * matrix[1] + v[1] * matrix[5] + v[2] * matrix[9] + matrix[13];
                var z = v[0] * matrix[2] + v[1] * matrix[6] + v[2] * matrix[10] + matrix[14];

                return Quantize(new Vector3((float)x, (float)y, (float)z), 1e-4f);
            }

            private static Vector3 GetNormal(COMApi.InwSimpleVertex vertex, double[] matrix)
            {
                try
                {
                    // InwSimpleVertex.normal доступен при eNORMAL
                    var arr = (Array)vertex.normal;
                    if (arr == null) return Vector3.Zero;

                    var v = arr.Cast<float>().ToArray();
                    if (v.Length < 3) return Vector3.Zero;

                    var n = new Vector3(v[0], v[1], v[2]);

                    if (matrix != null && matrix.Length == 16)
                    {
                        // Трансформируем нормаль 3x3 частью (без сдвига), затем нормализуем.
                        var x = n.X * matrix[0] + n.Y * matrix[4] + n.Z * matrix[8];
                        var y = n.X * matrix[1] + n.Y * matrix[5] + n.Z * matrix[9];
                        var z = n.X * matrix[2] + n.Y * matrix[6] + n.Z * matrix[10];
                        n = new Vector3((float)x, (float)y, (float)z);
                    }

                    if (n == Vector3.Zero) return Vector3.Zero;
                    n = Vector3.Normalize(n);
                    return Quantize(n, 1e-3f);
                }
                catch
                {
                    return Vector3.Zero;
                }
            }

            private static Vector3 Quantize(Vector3 v, float step)
            {
                if (step <= 0f) return v;
                float qx = (float)Math.Round(v.X / step) * step;
                float qy = (float)Math.Round(v.Y / step) * step;
                float qz = (float)Math.Round(v.Z / step) * step;
                return new Vector3(qx, qy, qz);
            }
        }

        // From http://adndevblog.typepad.com/aec/2012/08/geometry-fragment-returns-all-instances-when-a-multiply-instanced-node-is-selected.html
        private static bool IsFragmentOnPath(COMApi.InwOaPath3 path, COMApi.InwOaFragment3 fragment)
        {
            var a1 = (Array)fragment.path.ArrayData;
            var a2 = (Array)path.ArrayData;

            if (a1.GetLength(0) == a2.GetLength(0) &&
                a1.GetLowerBound(0) == a2.GetLowerBound(0) &&
                a1.GetUpperBound(0) == a2.GetUpperBound(0))
            {
                var i = a1.GetLowerBound(0);
                for (; i <= a1.GetUpperBound(0); i++)
                {
                    if ((int)a1.GetValue(i) != (int)a2.GetValue(i))
                        return false;
                }
            }

            return true;
        }

        private static double[] GetLocalToWorldTransformMatrix(COMApi.InwOaFragment3 fragment)
        {
            // Адаптация расширения из primer: GetLocalToWorldMatrix() возвращает InwLTransform3f3.
            var localToWorld = (COMApi.InwLTransform3f3)fragment.GetLocalToWorldMatrix();
            var matrix = (Array)localToWorld.Matrix;
            return matrix.Cast<double>().ToArray();
        }

        private static string SanitizeFileName(string name)
        {
            var invalid = Path.GetInvalidFileNameChars();
            var safe = new string(name.Select(c => invalid.Contains(c) ? '_' : c).ToArray());
            return string.IsNullOrWhiteSpace(safe) ? "Clash" : safe;
        }

        private static string BuildSafeNamedSegment(string prefix, string displayName, int maxSegmentLen)
        {
            // Делаем сегмент пути достаточно коротким, чтобы не упереться в лимит 260 символов на Windows.
            // Формат: "<prefix>_<sanitizedTrimmed>_<hash8>" (hash добавляем только если есть displayName).
            var sanitizedPrefix = SanitizeFileName(prefix ?? "Item");
            if (string.IsNullOrWhiteSpace(displayName))
                return TrimToLength(sanitizedPrefix, maxSegmentLen);

            var sanitizedName = SanitizeFileName(displayName).Trim();
            var hash8 = ShortHash8(displayName);

            // Оставляем место под "_"+hash
            var baseMax = Math.Max(1, maxSegmentLen);

            var combinedBase = sanitizedName;
            combinedBase = TrimToLength(combinedBase, baseMax);

            return combinedBase;
        }

        // Упрощённый вариант без явного префикса — используется в текущих вызовах
        private static string BuildSafeNamedSegment(string displayName, int maxSegmentLen)
        {
            return BuildSafeNamedSegment("Item", displayName, maxSegmentLen);
        }

        private static string TrimToLength(string value, int maxLen)
        {
            if (string.IsNullOrEmpty(value)) return value;
            if (maxLen <= 0) return string.Empty;
            return value.Length <= maxLen ? value : value.Substring(0, maxLen);
        }

        private static string ShortHash8(string value)
        {
            using (var sha1 = SHA1.Create())
            {
                var bytes = Encoding.UTF8.GetBytes(value ?? string.Empty);
                var hash = sha1.ComputeHash(bytes);
                // 8 hex символов достаточно для уникальности в рамках одного теста
                return BitConverter.ToString(hash, 0, 4).Replace("-", "").ToLowerInvariant();
            }
        }
    }
}

