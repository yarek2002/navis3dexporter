using System;
using System.IO;
using System.Reflection;

namespace Navis3dExporter
{
    internal static class AssemblyResolver
    {
        private static bool _registered;

        public static void EnsureRegistered()
        {
            if (_registered) return;
            _registered = true;

            AppDomain.CurrentDomain.AssemblyResolve += CurrentDomain_AssemblyResolve;
        }

        private static Assembly CurrentDomain_AssemblyResolve(object sender, ResolveEventArgs args)
        {
            try
            {
                var requested = new AssemblyName(args.Name).Name;
                if (string.IsNullOrWhiteSpace(requested)) return null;

                // Пытаемся грузить зависимости из папки, где лежит DLL плагина.
                var baseDir = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
                if (string.IsNullOrWhiteSpace(baseDir)) return null;

                var candidate = Path.Combine(baseDir, requested + ".dll");
                if (!File.Exists(candidate)) return null;

                return Assembly.LoadFrom(candidate);
            }
            catch
            {
                return null;
            }
        }
    }
}

