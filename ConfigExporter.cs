
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using NPOI.HSSF.UserModel; // For .xls files
using System;
using System.IO;
using System.Linq;
using Newtonsoft.Json;

namespace ExcelTable
{
    public class ConfigExporter
    {
        private readonly string _inputDirectory;
        private readonly string _outputDirectory;

        public ConfigExporter(string inputDirectory, string outputDirectory)
        {
            _inputDirectory = inputDirectory;
            _outputDirectory = outputDirectory;
        }

        /// <summary>
        /// 執行導出流程，處理所有Excel文件。
        /// </summary>
        public void ExportAll()
        {
            Console.WriteLine($"Starting export process...");
            Console.WriteLine($"Input directory: '{_inputDirectory}'");
            Console.WriteLine($"Output directory: '{_outputDirectory}'");

            // 1. 準備輸出目錄
            var schemaOutDir = Path.Combine(_outputDirectory, "Schemas");
            var jsonOutDir = Path.Combine(_outputDirectory, "Json");
            var xmlOutDir = Path.Combine(_outputDirectory, "Xml");

            Directory.CreateDirectory(schemaOutDir);
            Directory.CreateDirectory(jsonOutDir);
            Directory.CreateDirectory(xmlOutDir);

            // 2. 查找所有 Excel 文件 (xlsx 和 xls)
            var excelFiles = Directory.GetFiles(_inputDirectory, "*.xlsx", SearchOption.AllDirectories)
                .Concat(Directory.GetFiles(_inputDirectory, "*.xls", SearchOption.AllDirectories));

            foreach (var filePath in excelFiles)
            {
                Console.WriteLine($"\n--- Processing file: {Path.GetFileName(filePath)} ---");
                try
                {
                    ProcessWorkbook(filePath, schemaOutDir, jsonOutDir, xmlOutDir);
                }
                catch (Exception ex)
                {
                    Console.ForegroundColor = ConsoleColor.Red;
                    Console.WriteLine($"Failed to process file {filePath}. Error: {ex.Message}");
                    Console.ResetColor();
                }
            }

            Console.ForegroundColor = ConsoleColor.Green;
            Console.WriteLine("\nExport process completed successfully!");
            Console.ResetColor();
        }

        /// <summary>
        /// 處理單個Excel工作簿中的所有Sheet。
        /// </summary>
        private void ProcessWorkbook(string filePath, string schemaOutDir, string jsonOutDir, string xmlOutDir)
        {
            using var fileStream = new FileStream(filePath, FileMode.Open, FileAccess.Read);
            IWorkbook workbook = Path.GetExtension(filePath).ToLower() == ".xls"
                ? new HSSFWorkbook(fileStream)
                : new XSSFWorkbook(fileStream);

            for (int i = 0; i < workbook.NumberOfSheets; i++)
            {
                ISheet sheet = workbook.GetSheetAt(i);
                string sheetName = sheet.SheetName;

                // 3. 檢查Sheet名稱是否符合 "XXXX_Table" 規則
                if (sheetName.EndsWith("_Table", StringComparison.OrdinalIgnoreCase))
                {
                    // 4. 派生類名
                    string className = sheetName.Substring(0, sheetName.Length - "_Table".Length);
                    Console.WriteLine($"  -> Found matching sheet: '{sheetName}'. Generating class '{className}'...");

                    try
                    {
                        // 5. 使用 ExcelSheetParser 進行解析
                        var parser = new ExcelSheetParser(sheet, className);
                        parser.Parse();

                        // 6. 獲取結果並寫入文件
                        var schemaJson = JsonConvert.SerializeObject(parser.GetSchema(), Formatting.Indented);
                        File.WriteAllText(Path.Combine(schemaOutDir, $"{className}.schema.json"), schemaJson);

                        var jsonData = parser.GetJsonData();
                        File.WriteAllText(Path.Combine(jsonOutDir, $"{className}.json"), jsonData);

                        var xmlData = parser.GetXmlData();
                        File.WriteAllText(Path.Combine(xmlOutDir, $"{className}.xml"), xmlData);

                        Console.ForegroundColor = ConsoleColor.Green;
                        Console.WriteLine($"     Success: Generated Schema, JSON, and XML for '{className}'.");
                        Console.ResetColor();
                    }
                    catch (Exception ex)
                    {
                        Console.ForegroundColor = ConsoleColor.Yellow;
                        Console.WriteLine($"     Warning: Failed to parse sheet '{sheetName}'. Reason: {ex.Message}");
                        Console.ResetColor();
                    }
                }
            }
        }
    }
}