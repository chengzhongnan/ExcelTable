using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelTable
{
    /// <summary>
    /// 描述一个表格的整体结构
    /// </summary>
    public record TableSchema(string ClassName, List<PropertySchema> Properties);

    /// <summary>
    /// 描述表格中的一个属性（一列或一组列）
    /// </summary>
    public record PropertySchema(string Name, string Type, string? Comment, List<PropertySchema>? Children = null);

}
