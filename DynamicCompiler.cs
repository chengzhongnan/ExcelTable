using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.CSharp;
using Microsoft.CodeAnalysis.Emit;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelTable
{
    public static class DynamicCompiler
    {
        private static List<MetadataReference>? _dotNetStandardReferences;

        public static bool Compile(
            string assemblyName,
            string outputDllPath,
            Dictionary<string, string> sourceFiles,
            Dictionary<string, string> embeddedJsonData)
        {
            Console.WriteLine("Starting dynamic compilation targeting .NET Standard 2.1...");

            var syntaxTrees = sourceFiles.Select(file => CSharpSyntaxTree.ParseText(file.Value, path: file.Key, encoding: Encoding.UTF8));

            var references = GetNetStandardReferences();
            if (references == null)
            {
                return false;
            }

            var compilationOptions = new CSharpCompilationOptions(OutputKind.DynamicallyLinkedLibrary)
                .WithOptimizationLevel(OptimizationLevel.Release)
                .WithPlatform(Platform.AnyCpu);

            var compilation = CSharpCompilation.Create(assemblyName, syntaxTrees, references, compilationOptions);

            var manifestResources = new List<ResourceDescription>();
            foreach (var jsonData in embeddedJsonData)
            {
                string resourceName = $"{assemblyName}.{jsonData.Key}";
                byte[] resourceBytes = Encoding.UTF8.GetBytes(jsonData.Value);
                var resourceStreamProvider = new Func<Stream>(() => new MemoryStream(resourceBytes));
                var resource = new ResourceDescription(resourceName, resourceStreamProvider, isPublic: true);
                manifestResources.Add(resource);
            }

            string xmlDocPath = Path.ChangeExtension(outputDllPath, ".xml");

            using (var dllStream = new FileStream(outputDllPath, FileMode.Create))
            using (var pdbStream = new FileStream(Path.ChangeExtension(outputDllPath, ".pdb"), FileMode.Create))
            using (var xmlStream = new FileStream(xmlDocPath, FileMode.Create))
            {
                EmitResult result = compilation.Emit(
                    peStream: dllStream,
                    pdbStream: pdbStream,
                    xmlDocumentationStream: xmlStream,
                    manifestResources: manifestResources
                );

                if (!result.Success)
                {
                    Console.ForegroundColor = ConsoleColor.Red;
                    Console.WriteLine("Compilation failed!");
                    foreach (var diagnostic in result.Diagnostics.Where(d => d.Severity == DiagnosticSeverity.Error))
                    {
                        Console.WriteLine($"  {diagnostic.Id}: {diagnostic.GetMessage()} at {diagnostic.Location}");
                    }
                    Console.ResetColor();
                    dllStream.Close();
                    File.Delete(outputDllPath);
                    File.Delete(xmlDocPath);
                    return false;
                }
            }

            Console.ForegroundColor = ConsoleColor.Green;
            Console.WriteLine($"Successfully compiled '{Path.GetFileName(outputDllPath)}' (targeting .NET Standard 2.1).");
            Console.ResetColor();
            return true;
        }

        /// <summary>
        /// 【核心修改】採用更可靠的方式尋找 .NET Standard 2.1 的參考組件。
        /// </summary>
        private static List<MetadataReference>? GetNetStandardReferences()
        {
            if (_dotNetStandardReferences != null)
            {
                return _dotNetStandardReferences;
            }

            // 1. 尋找 dotnet.exe 的根目錄
            string? dotnetRoot = FindDotnetRootPath();
            if (string.IsNullOrEmpty(dotnetRoot))
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("Error: Could not find 'dotnet.exe' in the system's PATH. Please ensure the .NET SDK is installed correctly.");
                Console.ResetColor();
                return null;
            }

            // 2. 根據根目錄建構 packs 資料夾的路徑
            string packsPath = Path.Combine(dotnetRoot, "packs");
            string refAsmPath = Path.Combine(packsPath, "NETStandard.Library.Ref");

            if (!Directory.Exists(refAsmPath))
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine($"Error: .NET Standard reference assemblies pack not found at expected path: {refAsmPath}");
                Console.WriteLine("Please ensure the .NET SDK (not just the runtime) is installed.");
                Console.ResetColor();
                return null;
            }

            // 3. 找到最新版本的參考包
            string? latestVersionPath = Directory.GetDirectories(refAsmPath).OrderByDescending(d => d).FirstOrDefault();
            if (latestVersionPath == null)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine($"Error: No version folders found inside {refAsmPath}.");
                Console.ResetColor();
                return null;
            }

            string netStandard21Path = Path.Combine(latestVersionPath, "ref", "netstandard2.1");

            if (!Directory.Exists(netStandard21Path))
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine($"Error: .NET Standard 2.1 reference assemblies not found at path: {netStandard21Path}");
                Console.ResetColor();
                return null;
            }

            Console.WriteLine($"Loading .NET Standard 2.1 reference assemblies from: {netStandard21Path}");
            _dotNetStandardReferences = Directory.GetFiles(netStandard21Path, "*.dll")
                                                 .Select(dll => MetadataReference.CreateFromFile(dll))
                                                 .ToList<MetadataReference>();

            return _dotNetStandardReferences;
        }

        /// <summary>
        /// 透過搜尋系統 PATH 環境變數來找到 dotnet 的安裝根目錄。
        /// </summary>
        private static string? FindDotnetRootPath()
        {
            string pathVar = Environment.GetEnvironmentVariable("PATH");
            string executableName = Environment.OSVersion.Platform == PlatformID.Win32NT ? "dotnet.exe" : "dotnet";

            foreach (string path in pathVar.Split(Path.PathSeparator))
            {
                string fullPath = Path.Combine(path, executableName);
                if (File.Exists(fullPath))
                {
                    // dotnet.exe 所在的目錄就是 dotnet 的根目錄
                    return path;
                }
            }
            return null;
        }
    }
}
