using Newtonsoft.Json;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;

namespace ExcelTable
{

    public class ExcelSheetParser
    {
        private readonly ISheet _sheet;
        private readonly string _className;
        private int _headerRowCount;
        private int _typeRowIndex;
        private TableSchema? _schema;
        private List<Dictionary<string, object>>? _data;
        private readonly List<HeaderNode> _headerTree = new();

        // 用於解析表頭的中間數據結構
        private record HeaderNode(string Name, int StartCol, int EndCol, List<HeaderNode> Children);

        public ExcelSheetParser(ISheet sheet, string className)
        {
            _sheet = sheet;
            _className = className;
        }

        /// <summary>
        /// 執行完整的解析流程。
        /// </summary>
        public void Parse()
        {
            DetectHeaderRowCount();
            ParseHeader();
            ParseData();
        }

        /// <summary>
        /// 獲取解析後的表格結構 Schema。
        /// </summary>
        public TableSchema GetSchema()
        {
            if (_schema == null)
                throw new InvalidOperationException("Must call Parse() before getting the schema.");
            return _schema;
        }

        /// <summary>
        /// 以 JSON 字串格式獲取解析後的數據。
        /// </summary>
        public string GetJsonData(bool indented = true)
        {
            if (_data == null)
                throw new InvalidOperationException("Must call Parse() before getting the data.");
            return JsonConvert.SerializeObject(_data, indented ? Formatting.Indented : Formatting.None);
        }

        /// <summary>
        /// 以 XML 字串格式獲取解析後的數據。
        /// </summary>
        public string GetXmlData(bool indented = true)
        {
            if (_data == null)
                throw new InvalidOperationException("Must call Parse() before getting the data.");

            var root = new XElement(_className + "s");
            foreach (var rowData in _data)
            {
                var rowElement = new XElement(_className);
                BuildXmlElement(rowElement, rowData);
                root.Add(rowElement);
            }
            return root.ToString(indented ? SaveOptions.None : SaveOptions.DisableFormatting);
        }

        // =======================================================================
        // Private Helper Methods
        // =======================================================================

        #region Header Parsing

        /// <summary>
        /// 偵測表頭區域，找到類型定義行，並計算數據起始行。
        /// </summary>
        private void DetectHeaderRowCount()
        {
            for (int i = _sheet.FirstRowNum; i <= _sheet.LastRowNum; i++)
            {
                var row = _sheet.GetRow(i);
                if (row == null || row.LastCellNum <= 0) continue;

                int typeCellCount = 0;
                for (int j = row.FirstCellNum; j < row.LastCellNum; j++)
                {
                    var cellValue = GetStringCellValue(row.GetCell(j))?.ToLowerInvariant();
                    if (cellValue is "int" or "string" or "bool" || cellValue?.StartsWith("list") == true)
                    {
                        typeCellCount++;
                    }
                }

                if (typeCellCount > row.LastCellNum / 2)
                {
                    _typeRowIndex = i;
                    // 數據行在 類型行(i) 和 註釋行(i+1) 之後，所以是 i + 2
                    _headerRowCount = i + 2;
                    return;
                }
            }
            throw new Exception($"Could not detect the header type row in sheet '{_sheet.SheetName}'. A row with types like 'int', 'string' is required.");
        }

        /// <summary>
        /// 解析表頭，生成 Schema。
        /// </summary>
        private void ParseHeader()
        {
            var firstRow = _sheet.GetRow(0);
            if (firstRow == null) throw new Exception($"Sheet '{_sheet.SheetName}' is empty or header is missing.");

            BuildHeaderTreeRecursive(0, firstRow.FirstCellNum, firstRow.LastCellNum - 1, _headerTree);
            var properties = _headerTree.Select(ConvertToSchema).ToList();
            _schema = new TableSchema(_className, properties);
        }

        /// <summary>
        /// 遞迴地建立表頭的樹狀結構。
        /// </summary>
        private void BuildHeaderTreeRecursive(int rowIndex, int startCol, int endCol, List<HeaderNode> parentChildren)
        {
            if (rowIndex >= _typeRowIndex) return;

            var row = _sheet.GetRow(rowIndex);
            if (row == null) return;

            var processedCols = new HashSet<int>();
            for (int i = startCol; i <= endCol; i++)
            {
                if (processedCols.Contains(i)) continue;

                var cell = row.GetCell(i);
                var mergedRegion = GetMergedRegionForCell(cell);

                int colSpanStart = i;
                int colSpanEnd = i;
                if (mergedRegion != null && mergedRegion.FirstRow == rowIndex)
                {
                    colSpanStart = mergedRegion.FirstColumn;
                    colSpanEnd = mergedRegion.LastColumn;
                }

                var name = GetStringCellValue(cell);
                if (string.IsNullOrEmpty(name)) { processedCols.Add(i); continue; }

                var node = new HeaderNode(name, colSpanStart, colSpanEnd, new List<HeaderNode>());
                parentChildren.Add(node);

                BuildHeaderTreeRecursive(rowIndex + 1, colSpanStart, colSpanEnd, node.Children);

                for (int j = colSpanStart; j <= colSpanEnd; j++) { processedCols.Add(j); }
            }
        }

        /// <summary>
        /// 將中間的 HeaderNode 樹轉換為最終的 PropertySchema 結構，包含智能判斷邏輯。
        /// </summary>
        private PropertySchema ConvertToSchema(HeaderNode node)
        {
            var commentRow = _sheet.GetRow(_typeRowIndex + 1);
            string? comment = commentRow?.GetCell(node.StartCol)?.ToString()?.Trim();

            // 規則A: 如果一個節點只有一個子節點，且該子節點無後代 (通常是簡單欄位帶一個描述行)，則視為簡單類型
            if (node.Children.Count == 1 && node.Children[0].Children.Count == 0)
            {
                var type = GetStringCellValue(_sheet.GetRow(_typeRowIndex).GetCell(node.StartCol)) ?? "string";
                return new PropertySchema(node.Name, type, comment);
            }

            // 規則B: 如果節點無子節點，必然是簡單類型
            if (node.Children.Count == 0)
            {
                var type = GetStringCellValue(_sheet.GetRow(_typeRowIndex).GetCell(node.StartCol)) ?? "string";
                return new PropertySchema(node.Name, type, comment);
            }

            var firstChild = node.Children[0];

            // 規則C: 簡單類型列表 (e.g., LiveScene -> Scene, Scene, Scene)
            if (node.Children.All(c => c.Name == firstChild.Name && c.Children.Count == 0))
            {
                var itemType = GetStringCellValue(_sheet.GetRow(_typeRowIndex).GetCell(firstChild.StartCol)) ?? "string";
                return new PropertySchema(node.Name, $"List<{itemType}>", comment);
            }

            // 規則D: 物件列表 (e.g., Skills -> Name, Level, Name, Level)
            int patternLength = node.Children.FindIndex(1, c => c.Name == firstChild.Name);
            if (patternLength > 0 && node.Children.Count % patternLength == 0)
            {
                var itemSchemaChildren = node.Children.Take(patternLength).Select(ConvertToSchema).ToList();
                var objectTypeName = $"{node.Name}Object";
                return new PropertySchema(node.Name, $"List<{objectTypeName}>", comment, itemSchemaChildren);
            }

            // 規則E: 單一的嵌套物件 (e.g., BaseProperty -> Hp, Attack)
            var objectChildren = node.Children.Select(ConvertToSchema).ToList();
            return new PropertySchema(node.Name, $"{node.Name}Object", comment, objectChildren);
        }
        #endregion

        #region Data Parsing

        /// <summary>
        /// 解析所有數據行，生成結構化的數據列表。
        /// </summary>
        private void ParseData()
        {
            _data = new List<Dictionary<string, object>>();
            if (_schema == null) return;

            for (int i = _headerRowCount; i <= _sheet.LastRowNum; i++)
            {
                var row = _sheet.GetRow(i);
                if (row == null || row.GetCell(row.FirstCellNum) == null || string.IsNullOrWhiteSpace(row.GetCell(row.FirstCellNum).ToString()))
                {
                    continue;
                }

                int currentCol = 0;
                var rowData = new Dictionary<string, object>();
                foreach (var prop in _schema.Properties)
                {
                    rowData[prop.Name] = ParsePropertyData(row, prop, ref currentCol);
                }
                _data.Add(rowData);
            }
        }

        /// <summary>
        /// 遞迴地解析單個屬性的數據，具備類型感知能力。
        /// </summary>
        private object? ParsePropertyData(IRow row, PropertySchema prop, ref int currentCol)
        {
            if (prop.Type.StartsWith("List<"))
            {
                var list = new List<object>();
                string itemType = prop.Type.Substring(5, prop.Type.Length - 6);

                if (prop.Children != null) // List of Objects
                {
                    int colSpan = prop.Children.Count;
                    int totalSpan = GetColumnCountFromHeader(prop.Name);
                    int startCol = currentCol;
                    currentCol += totalSpan;

                    for (int i = startCol; i < startCol + totalSpan; i += colSpan)
                    {
                        if (row.GetCell(i) == null || string.IsNullOrWhiteSpace(row.GetCell(i).ToString())) break;
                        var itemObj = new Dictionary<string, object>();
                        int tempCol = i;
                        foreach (var childProp in prop.Children)
                        {
                            itemObj[childProp.Name] = ParsePropertyData(row, childProp, ref tempCol);
                        }
                        list.Add(itemObj);
                    }
                }
                else // List of Primitives
                {
                    int totalSpan = GetColumnCountFromHeader(prop.Name);
                    int startCol = currentCol;
                    currentCol += totalSpan;
                    for (int i = startCol; i < startCol + totalSpan; i++)
                    {
                        var cell = row.GetCell(i);
                        if (cell == null || string.IsNullOrWhiteSpace(cell.ToString())) break;
                        list.Add(GetTypedCellValue(cell, itemType)!);
                    }
                }
                return list;
            }
            else if (prop.Children != null) // Single Object
            {
                var obj = new Dictionary<string, object>();
                foreach (var childProp in prop.Children)
                {
                    obj[childProp.Name] = ParsePropertyData(row, childProp, ref currentCol);
                }
                return obj;
            }
            else // Simple Type
            {
                var cell = row.GetCell(currentCol);
                currentCol++;
                return GetTypedCellValue(cell, prop.Type);
            }
        }

        /// <summary>
        /// 根據指定的 C# 類型，獲取單元格的值。
        /// </summary>
        private object? GetTypedCellValue(ICell? cell, string csharpType)
        {
            if (cell == null) return null;

            if (cell.CellType == CellType.Formula)
            {
                try { cell = cell.Sheet.Workbook.GetCreationHelper().CreateFormulaEvaluator().EvaluateInCell(cell); }
                catch { /* fallback to string */ }
            }

            switch (csharpType.ToLower())
            {
                case "int":
                    if (cell.CellType == CellType.Numeric) return Convert.ToInt32(cell.NumericCellValue);
                    if (int.TryParse(cell.ToString(), out int intValue)) return intValue;
                    return 0;
                case "float":
                    if (cell.CellType == CellType.Numeric) return Convert.ToSingle(cell.NumericCellValue);
                    if (float.TryParse(cell.ToString(), out float floatValue)) return floatValue;
                    return 0f;
                case "double":
                    if (cell.CellType == CellType.Numeric) return cell.NumericCellValue;
                    if (double.TryParse(cell.ToString(), out double doubleValue)) return doubleValue;
                    return 0.0;
                case "bool":
                    if (cell.CellType == CellType.Boolean) return cell.BooleanCellValue;
                    return cell.ToString()?.ToLower() == "true" || cell.ToString() == "1";
                case "string":
                default:
                    return cell.ToString()?.Trim();
            }
        }
        #endregion

        #region Misc Helpers
        private string? GetStringCellValue(ICell? cell)
        {
            if (cell == null) return null;
            if (cell.CellType == CellType.Formula)
            {
                try { return cell.Sheet.Workbook.GetCreationHelper().CreateFormulaEvaluator().Evaluate(cell).FormatAsString()?.Trim(); }
                catch { return cell.ToString()?.Trim(); }
            }
            return cell.ToString()?.Trim();
        }

        private int GetColumnCountFromHeader(string propertyName)
        {
            // 查找頂層節點
            var node = _headerTree.Find(n => n.Name == propertyName);
            if (node != null)
            {
                return node.EndCol - node.StartCol + 1;
            }
            // 如果在頂層找不到，可能是一個嵌套屬性，這裡簡化處理
            return 1;
        }

        private CellRangeAddress? GetMergedRegionForCell(ICell cell)
        {
            if (cell == null) return null;
            return _sheet.MergedRegions.FirstOrDefault(r => r.IsInRange(cell.RowIndex, cell.ColumnIndex));
        }

        private void BuildXmlElement(XElement parent, object data)
        {
            if (data is Dictionary<string, object> dict) { foreach (var kvp in dict) { var childElement = new XElement(kvp.Key); BuildXmlElement(childElement, kvp.Value); parent.Add(childElement); } }
            else if (data is List<object> list) { foreach (var item in list) { var itemElement = new XElement("Item"); BuildXmlElement(itemElement, item); parent.Add(itemElement); } }
            else { parent.Value = data?.ToString() ?? ""; }
        }
        #endregion
    }

}
