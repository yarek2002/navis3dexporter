using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Numerics;
using System.Security.Cryptography;
using System.Text;
using Autodesk.Navisworks.Api;
using Autodesk.Navisworks.Api.Clash;
using SharpGLTF.Geometry;
using SharpGLTF.Geometry.VertexTypes;
using SharpGLTF.Materials;
using SharpGLTF.Scenes;
using ComBridge = Autodesk.Navisworks.Api.ComApi.ComApiBridge;
using COMApi = Autodesk.Navisworks.Api.Interop.ComApi;

namespace Navis3dExporter
{
    internal class GlbClashExporter
    {
        private readonly Document _document;

        public GlbClashExporter(Document document)
        {
            _document = document ?? throw new ArgumentNullException(nameof(document));
        }

        public void ExportCurrentSelection(string outputFilePath)
        {
            if (string.IsNullOrWhiteSpace(outputFilePath))
                throw new ArgumentException("Output file path is not specified.", nameof(outputFilePath));

            var selection = _document.CurrentSelection;
            if (selection == null || selection.SelectedItems == null || selection.SelectedItems.Count == 0)
                throw new InvalidOperationException("В текущем выделении нет элементов для экспорта.");

            var scene = new SceneBuilder();

            int index = 0;
            foreach (var item in selection.SelectedItems)
            {
                index++;

                // Чередуем цвета, чтобы элементы различались
                var color = (index % 2 == 1)
                    ? new Vector4(1f, 0f, 0f, 1f)
                    : new Vector4(0f, 0f, 1f, 1f);

                var mesh = BuildMeshFromModelItem(item, $"Sel_{index}", color);
                if (mesh != null)
                {
                    scene.AddRigidMesh(mesh, Matrix4x4.Identity);
                }
            }

            var model = scene.ToGltf2();
            model.SaveGLB(outputFilePath);
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
                            displayName: group.DisplayName,
                            maxSegmentLen: 80);

                        string filePath = Path.Combine(testFolder, groupName + ".glb");
                        ExportGroup(group, filePath);
                    }
                    else if (child is ClashResult clashResult)
                    {
                        // Одиночные результаты без группы
                        clashIndex++;
                        string clashName = BuildSafeNamedSegment(
                            displayName: clashResult.DisplayName,
                            maxSegmentLen: 80);

                        string filePath = Path.Combine(testFolder, clashName + ".glb");
                        ExportSingleClash(clashResult, filePath);
                    }
                }
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
                scene.AddRigidMesh(mesh1, Matrix4x4.Identity);

            // Второй элемент – синим
            var mesh2 = BuildMeshFromModelItem(item2, "Item2",
                new Vector4(0f, 0f, 1f, 1f));
            if (mesh2 != null)
                scene.AddRigidMesh(mesh2, Matrix4x4.Identity);

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
                    scene.AddRigidMesh(mesh1, Matrix4x4.Identity);

                var mesh2 = BuildMeshFromModelItem(item2, $"Item2_{index}",
                    new Vector4(0f, 0f, 1f, 1f));
                if (mesh2 != null)
                    scene.AddRigidMesh(mesh2, Matrix4x4.Identity);
            }

            var model = scene.ToGltf2();
            model.SaveGLB(filePath);
        }

        private MeshBuilder<VertexPositionNormal, VertexEmpty, VertexEmpty> BuildMeshFromModelItem(
            ModelItem modelItem,
            string meshName,
            Vector4 color)
        {
           var triangles = ExtractTriangles(modelItem);
            System.Windows.MessageBox.Show(
                $"Triangles for {meshName}: {triangles.Count}",
                "GLB Exporter");
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
                var v0 = new VertexPositionNormal(tri.V0, tri.Normal);
                var v1 = new VertexPositionNormal(tri.V1, tri.Normal);
                var v2 = new VertexPositionNormal(tri.V2, tri.Normal);

                prim.AddTriangle(v0, v1, v2);
            }

            return mesh;
        }

        private struct TriangleData
        {
            public Vector3 V0;
            public Vector3 V1;
            public Vector3 V2;
            public Vector3 Normal;
        }

        private List<TriangleData> ExtractTriangles(ModelItem modelItem)
        {
            var triangles = new List<TriangleData>();

            // Берём сам элемент и всю его иерархию, как в Clash Detective
            var coll = new ModelItemCollection();
            coll.AddRange(modelItem.DescendantsAndSelf);

            var selection = (COMApi.InwOpSelection)ComBridge.ToInwOpSelection(coll);

            var callback = new TriangleCollector(triangles);

            foreach (COMApi.InwOaPath3 path in selection.Paths())
            {
                foreach (COMApi.InwOaFragment3 frag in path.Fragments())
                {
                    callback.CurrentTransform = TryGetFragmentTransform(frag);

                    frag.GenerateSimplePrimitives(
                        COMApi.nwEVertexProperty.eNORMAL,
                        callback);
                }
            }

            return triangles;
        }

        private class TriangleCollector : COMApi.InwSimplePrimitivesCB
        {
            private readonly List<TriangleData> _triangles;

            public Matrix4x4 CurrentTransform { get; set; } = Matrix4x4.Identity;

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
                var p0 = Vector3.Transform(ToVector3(v1), CurrentTransform);
                var p1 = Vector3.Transform(ToVector3(v2), CurrentTransform);
                var p2 = Vector3.Transform(ToVector3(v3), CurrentTransform);

                var normal = Vector3.Normalize(Vector3.Cross(p1 - p0, p2 - p0));

                _triangles.Add(new TriangleData
                {
                    V0 = p0,
                    V1 = p1,
                    V2 = p2,
                    Normal = normal
                });
            }

            private static Vector3 ToVector3(COMApi.InwSimpleVertex v)
            {
                var coord = (Array)v.coord;
                return new Vector3(
                    (float)(double)coord.GetValue(0),
                    (float)(double)coord.GetValue(1),
                    (float)(double)coord.GetValue(2));
            }
        }

        private static Matrix4x4 TryGetFragmentTransform(COMApi.InwOaFragment3 frag)
        {
            try
            {
                // Возвращает variant array[16] (double) - матрица LCS->WCS
                var arr = (Array)frag.GetLocalToWorldMatrix();
                if (arr == null || arr.Length < 16) return Matrix4x4.Identity;

                float m00 = (float)(double)arr.GetValue(0);
                float m01 = (float)(double)arr.GetValue(1);
                float m02 = (float)(double)arr.GetValue(2);
                float m03 = (float)(double)arr.GetValue(3);

                float m10 = (float)(double)arr.GetValue(4);
                float m11 = (float)(double)arr.GetValue(5);
                float m12 = (float)(double)arr.GetValue(6);
                float m13 = (float)(double)arr.GetValue(7);

                float m20 = (float)(double)arr.GetValue(8);
                float m21 = (float)(double)arr.GetValue(9);
                float m22 = (float)(double)arr.GetValue(10);
                float m23 = (float)(double)arr.GetValue(11);

                float m30 = (float)(double)arr.GetValue(12);
                float m31 = (float)(double)arr.GetValue(13);
                float m32 = (float)(double)arr.GetValue(14);
                float m33 = (float)(double)arr.GetValue(15);

                return new Matrix4x4(
                    m00, m01, m02, m03,
                    m10, m11, m12, m13,
                    m20, m21, m22, m23,
                    m30, m31, m32, m33);
            }
            catch
            {
                return Matrix4x4.Identity;
            }
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
            var suffix = "_" + hash8;
            var baseMax = Math.Max(1, maxSegmentLen - suffix.Length);

            var combinedBase = sanitizedName;
            combinedBase = TrimToLength(combinedBase, baseMax);

            return combinedBase + suffix;
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

