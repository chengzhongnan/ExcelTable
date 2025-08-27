using Newtonsoft.Json;
using System.Xml.Serialization;

namespace ExcelTable
{
    internal class Program
    {
        static void Main(string[] args)
        {
            // =======================================================================
            // 階段 1: 從 AppConfig.xml 載入設定
            // =======================================================================
            string configFilePath = "AppConfig.xml";
            if (!File.Exists(configFilePath))
            {
                Console.ForegroundColor = ConsoleColor.Yellow;
                Console.WriteLine($"Configuration file '{configFilePath}' not found.");
                ConfigLoader.CreateDefault(configFilePath);
                Console.WriteLine($"A default '{configFilePath}' has been created. Please configure it and run the application again.");
                Console.ResetColor();
                return;
            }

            Console.WriteLine($"Loading settings from '{configFilePath}'...");
            AppSettings settings = ConfigLoader.Load(configFilePath);

            var inputDirectory = settings.InputPath;
            var outputDirectory = settings.OutputPath;
            var assemblyName = settings.AssemblyName;

            Console.WriteLine($"Input Path: {inputDirectory}");
            Console.WriteLine($"Output Path: {outputDirectory}");
            Console.WriteLine($"Assembly Name: {assemblyName}");

            // =======================================================================
            // 階段 2: 執行匯出和編譯流程 (後續邏輯不變，僅使用來自設定的變數)
            // =======================================================================

            if (!Directory.Exists(inputDirectory))
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine($"Input directory not found: '{inputDirectory}'");
                Console.ResetColor();
                return;
            }

            // 為了保持您原有的結構，我們仍然可以這樣調用 ConfigExporter
            // 或者，您可以將 ConfigExporter 的邏輯直接整合到這裡
            var exporter = new ConfigExporter(inputDirectory, outputDirectory);
            exporter.ExportAll();

            Console.WriteLine("\nReading generated files for compilation...");

            var allSchemas = new List<TableSchema>();
            var allJsonData = new Dictionary<string, string>();

            var schemaOutDir = Path.Combine(outputDirectory, "Schemas");
            var jsonOutDir = Path.Combine(outputDirectory, "Json");

            if (!Directory.Exists(schemaOutDir) || !Directory.Exists(jsonOutDir))
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("Export step did not create Schema/Json directories. Aborting compilation.");
                Console.ResetColor();
                return;
            }

            var schemaFiles = Directory.GetFiles(schemaOutDir, "*.schema.json");
            foreach (var schemaFile in schemaFiles)
            {
                var schemaJson = File.ReadAllText(schemaFile);
                var schema = JsonConvert.DeserializeObject<TableSchema>(schemaJson);
                if (schema != null) { allSchemas.Add(schema); }
            }
            Console.WriteLine($"Loaded {allSchemas.Count} schema files.");

            var jsonDataFiles = Directory.GetFiles(jsonOutDir, "*.json");
            foreach (var dataFile in jsonDataFiles)
            {
                var jsonData = File.ReadAllText(dataFile);
                var fileName = Path.GetFileName(dataFile);
                allJsonData.Add(fileName, jsonData);
            }
            Console.WriteLine($"Loaded {allJsonData.Count} json data files.");

            if (allSchemas.Count == 0)
            {
                Console.WriteLine("\nNo valid schemas found to process. Exiting.");
                return;
            }

            Console.WriteLine("\nGenerating C# source code in memory...");

            Dictionary<string, string> sourceFiles = CodeGenerator.GenerateAllSourceFiles(allSchemas, assemblyName);
            Console.WriteLine($"Generated {sourceFiles.Count} source files for namespace '{assemblyName}'.");

            string dllOutputDir = Path.Combine(outputDirectory, "DLL");
            Directory.CreateDirectory(dllOutputDir);
            string outputDllPath = Path.Combine(dllOutputDir, $"{assemblyName}.dll");

            bool success = DynamicCompiler.Compile(assemblyName, outputDllPath, sourceFiles, allJsonData);

            if (success)
            {
                Console.ForegroundColor = ConsoleColor.Cyan;
                Console.WriteLine($"\nDLL generation complete. Output file: {outputDllPath}");
                Console.ResetColor();
            }
            else
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("\nDLL generation failed. Please check the errors above.");
                Console.ResetColor();
            }
        }
    }
}
