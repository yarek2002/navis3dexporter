using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Numerics;
using Autodesk.Navisworks.Api;
using Autodesk.Navisworks.Api.Clash;
using SharpGLTF.Geometry;
using SharpGLTF.Geometry.VertexTypes;
using SharpGLTF.Materials;
using SharpGLTF.Scenes;
using COMApi = Autodesk.Navisworks.Api.Interop.ComApi;
using ComBridge = Autodesk.Navisworks.Api.ComApi.ComApiBridge;

namespace Navis3dExporter
{
    internal class GlbClashExporter
    {
        private readonly Document _document;

        public GlbClashExporter(Document document)
        {
            _document = document ?? throw new ArgumentNullException(nameof(document));
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
            int clashIndex = 0;

            foreach (var test in tests)
            {
                var clashTest = test as ClashTest;
                if (clashTest == null)
                    continue;

                var results = GetAllResults(clashTest);
                foreach (var result in results)
                {
                    clashIndex++;
                    string clashName = string.IsNullOrWhiteSpace(result.DisplayName)
                        ? $"Clash_{clashIndex:0000}"
                        : SanitizeFileName(result.DisplayName);

                    string filePath = Path.Combine(outputFolder, clashName + ".glb");
                    ExportSingleClash(result, filePath);
                }
            }
        }

        private static IEnumerable<ClashResult> GetAllResults(ClashTest clashTest)
        {
            // Разворачиваем и группы, и одиночные результаты
            var stack = new Stack<SavedItem>();
            foreach (var child in clashTest.Children)
                stack.Push(child);

            while (stack.Count > 0)
            {
                var item = stack.Pop();
                if (item is ClashResult clashResult)
                {
                    yield return clashResult;
                }
                else if (item is ClashResultGroup group)
                {
                    foreach (var child in group.Children)
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

            var comState = (COMApi.InwOpState)ComBridge.State;
            var comItem = ComBridge.ToInwOaPath(modelItem);
            var selection = comState.ObjectFactory(
                COMApi.nwEObjectType.eObjectType_nwOpSelection,
                null, null) as COMApi.InwOpSelection;

            selection.Select(comItem);

            var callback = new TriangleCollector(triangles);

            foreach (COMApi.InwOaPath3 path in selection.Paths())
            {
                foreach (COMApi.InwOaFragment3 frag in path.Fragments())
                {
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
                var p0 = ToVector3(v1);
                var p1 = ToVector3(v2);
                var p2 = ToVector3(v3);

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

        private static string SanitizeFileName(string name)
        {
            var invalid = Path.GetInvalidFileNameChars();
            var safe = new string(name.Select(c => invalid.Contains(c) ? '_' : c).ToArray());
            return string.IsNullOrWhiteSpace(safe) ? "Clash" : safe;
        }
    }
}

