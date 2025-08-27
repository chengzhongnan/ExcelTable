# **Excel 配置编译工具 (Excel to DLL Configuration Compiler)**

## **项目简介**

本项目是一个强大的 .NET 命令行工具，旨在将复杂的 Excel 配置文件（.xlsx/.xls）直接编译成一个类型安全、易于使用、零依赖的 **.NET Standard 2.1 DLL**。

在游戏开发和需要频繁调整配置的应用程序中，传统的手动复制粘贴或编写解析代码的方式效率低下且容易出错。本工具将整个流程自动化，让开发者可以专注于业务逻辑，同时为策划和运营人员提供他们熟悉的 Excel 作为配置工具。

## **功能特性**

* **动态解析多级表头**：能准确识别复杂的 Excel 表头结构，包括合并单元格代表的嵌套对象和列表。  
* **自动生成 C\# 类**：根据表头结构自动生成对应的强类型 C\# 数据类。  
* **类型安全**：生成的 DataManager 提供类型安全的数据访问接口，杜绝因手误写错字段名或类型转换错误导致的 Bug。  
* **编译为 DLL**：将生成的 C\# 类和 JSON 数据直接编译成一个独立的 .NET Standard 2.1 DLL，方便在任何 .NET 项目中引用。  
* **嵌入 JSON 数据**：在编译时，将 Excel 中的数据转换为 JSON 并作为内嵌资源打包进 DLL，实现数据的自包含。  
* **生成 XML 文件注释**：将 Excel 表头中的注释行编译为 XML 文件注释，在 Visual Studio 中提供完整的 IntelliSense 智能提示。  
* **灵活的数据加载**：生成的 DataManager 支持从 DLL 内嵌资源或外部文件夹两种模式加载配置，方便开发调试和部署。  
* **依赖注入解析**：DataManager 与具体的 JSON 解析库解耦，允许用户在主项目中提供自己的解析回调函数。  
* **外部化配置**：工具本身的所有路径和设置均由 AppConfig.xml 控制，无需修改代码。

## **如何使用**

#### **1\. 环境要求**

* 安装 [.NET SDK](https://dotnet.microsoft.com/download) (推荐 .NET 8.0 或更高版本)。

#### **2\. 设置 AppConfig.xml**

在工具的执行目录下，创建并设置 AppConfig.xml：

```
<?xml version="1.0" encoding="utf-8" ?>  
<Configuration>  
  <!-- 推荐使用相对于此工具 .exe 的相对路径 -->  
  <Paths>  
    <!-- 包含所有 .xlsx 文件的输入文件夹 -->  
    <InputPath>xlsx</InputPath>  
    <!-- 所有生成文件的根输出文件夹 -->  
    <OutputPath>output</OutputPath>  
  </Paths>  
  <Settings>  
    <!-- 生成的 DLL 的程序集名称，也将是 C# 类的根命名空间 -->  
    <AssemblyName>MyGame.Configs</AssemblyName>  
  </Settings>  
</Configuration>

```

#### **3\. Excel 表格规范**

为了让工具能正确解析，您的 Excel 表格需要遵循以下**四行表头**结构：

| 第 N 行 | 目的 | 示例 |
| :---- | :---- | :---- |
| **第 1 行** | **结构定义 (上)** | Id, Name, BaseProperty (合并单元格), Skills (合并单元格) |
| **第 2 行** | **结构定义 (下)** | (Id下为空), (Name下为空), Hp, Attack, Name, Level, Name, Level |
| **第 3 行** | **类型定义** | int, string, int, int, string, int, string, int |
| **第 4 行** | **注释/说明** | 唯一ID, 名字, 生命值, 攻击力, 技能名称, 技能等级, 技能名称, 技能等级 |
| **第 5 行** | **数据开始** | 1001, 小火龙, 100, 10, 火花, 1, 抓, 1 |

* **单一值**：表头只占一列 (如 Id, Name)。  
* **嵌套对象**：表头为合并单元格，其下方定义了对象的各个属性 (如 BaseProperty \-\> Hp, Attack)。  
* **对象列表**：表头为合并单元格，其下方是重复的对象结构 (如 Skills \-\> Name, Level, Name, Level...)。  
* **简单类型列表**：表头为合并单元格，其下方是重复的同名简单属性 (如 LiveScene \-\> Scene, Scene...)。

#### **4\. 运行工具**

将您的 Excel 文件（Sheet 页名称需以 \_Table 结尾，如 Boss\_Table.xlsx）放入 AppConfig.xml 中指定的 InputPath 文件夹内。  
然后执行工具：  
dotnet run

或直接执行编译好的 .exe 文件。

工具将在 OutputPath 中生成 DLL、Schemas 和 Json 文件夹。

#### **5\. 在项目中使用生成的 DLL**

1. **引用 DLL**：将 output/DLL/ 目录下的 MyGame.Configs.dll 和 MyGame.Configs.xml 复制到您的主项目中，并添加对 DLL 的引用。  
2. **编写加载代码**：在您的主项目中，使用 DeserializerFactory 辅助类（可从工具的异常信息中复制，或预先加入项目）来加载数据。  
``` 
   using MyGame.Configs;  
   using System.Reflection;

   public class GameInitializer  
   {  
       public void Initialize()  
       {  
           // 获取配置 DLL 的程序集  
           var configAssembly = typeof(DataManager).Assembly;

           // 使用工厂方法自动创建所有表格的 JSON 解析器  
           var deserializers = DeserializerFactory.CreateJsonDeserializersForAllTables(configAssembly, "MyGame.Configs");

           // 方式一：从 DLL 内嵌资源加载 (用于正式发布)  
           DataManager.LoadAllFromEmbedded(configAssembly, deserializers);

           // 方式二：从外部文件夹加载 (用于开发调试)  
           // string configPath = "path/to/your/json/files";  
           // DataManager.LoadAllFromDirectory(configPath, deserializers, ".json");

           // --- 加载完成，可以安全使用了 ---  
           var boss = DataManager.GetBossById(1001);  
           if (boss != null)  
           {  
               Console.WriteLine($"Boss Name: {boss.Name}, HP: {boss.BaseProperty.Hp}");  
           }  
       }

       public static Dictionary<Type, Func<Stream, IEnumerable<object>>> CreateJsonDeserializersForAllTables(Assembly configAssembly, string targetNamespace)
       {
            var deserializers = new Dictionary<Type, Func<Stream, IEnumerable<object>>>();

            var configTypes = configAssembly.GetTypes().Where(t =>
                t.IsClass &&
                t.IsPublic &&
                t.Namespace == targetNamespace &&
                t.Name != "DataManager"
            );

            foreach (var configType in configTypes)
            {
                deserializers[configType] = (stream) =>
                {
                    using var reader = new StreamReader(stream);
                    string json = reader.ReadToEnd();
                    Type listType = typeof(List<>).MakeGenericType(configType);
                    return (IEnumerable<object>)JsonConvert.DeserializeObject(json, listType);
                };
            }
            return deserializers;
       }
   }

```
## **项目依赖**

* [NPOI](https://www.google.com/search?q=https://github.com/nissim-aj/npoi) \- 用于读取 Excel 文件。  
* [Newtonsoft.Json](https://www.newtonsoft.com/json) \- 用于数据序列化和辅助类。  
* [Microsoft.CodeAnalysis.CSharp](https://www.nuget.org/packages/Microsoft.CodeAnalysis.CSharp/) (Roslyn) \- 用于动态编译 C\# 代码。

## **许可证**

本项目采用 [MIT](https://opensource.org/licenses/MIT) 许可证。