# **Excel to DLL Configuration Compiler**

## **About The Project**

This is a powerful .NET command-line tool designed to compile complex Excel configuration files (.xlsx/.xls) directly into a type-safe, easy-to-use, and dependency-free **.NET Standard 2.1 DLL**.

In game development and other applications requiring frequent configuration adjustments, traditional methods like manual copy-pasting or writing boilerplate parsing code are inefficient and error-prone. This tool automates the entire workflow, allowing developers to focus on business logic while providing designers and operators with the familiar Excel interface for managing configurations.

## **Features**

* **Dynamic Multi-Level Header Parsing**: Accurately interprets complex Excel header structures, including merged cells representing nested objects and lists.  
* **Automatic C\# Class Generation**: Automatically generates strongly-typed C\# data classes based on the header structure.  
* **Type-Safe Data Access**: The generated DataManager provides a type-safe API, eliminating bugs caused by typos in field names or incorrect type conversions.  
* **Compile to DLL**: Compiles the generated C\# classes and JSON data into a standalone .NET Standard 2.1 DLL for easy referencing in any .NET project.  
* **Embedded JSON Data**: Converts Excel data to JSON and embeds it as a resource into the DLL during compilation, making the data self-contained.  
* **XML Documentation Generation**: Compiles comment rows from the Excel header into XML documentation, providing full IntelliSense support in Visual Studio.  
* **Flexible Data Loading**: The generated DataManager supports loading configurations from either embedded resources or an external directory, facilitating both development and deployment.  
* **Dependency-Injected Deserialization**: The DataManager is decoupled from any specific JSON parsing library, allowing the user to provide their own deserialization callbacks.  
* **External Configuration**: All paths and settings for the tool itself are managed via an AppConfig.xml file, requiring no code changes for different environments.

## **How To Use**

#### **1\. Prerequisites**

* [.NET SDK](https://dotnet.microsoft.com/download) ( .NET 8.0 or higher recommended).

#### **2\. Configure AppConfig.xml**

In the tool's execution directory, create and configure AppConfig.xml:

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

#### **3\. Excel File Specification**

* Sheet Naming: For a sheet to be processed, its name must end with _Table (e.g., Boss_Table).

* Header Structure: The sheet must follow the four-row header structure below:

| Row \# | Purpose | Example |
| :---- | :---- | :---- |
| **Row 1** | **Structure (Top)** | Id, Name, BaseProperty (merged), Skills (merged) |
| **Row 2** | **Structure (Bottom)** | (blank), (blank), Hp, Attack, Name, Level, Name, Level |
| **Row 3** | **Type Definition** | int, string, int, int, string, int, string, int |
| **Row 4** | **Comment/Description** | Unique ID, Character Name, Health Points, Attack Power, Skill Name, Skill Level, ... |
| **Row 5** | **Data Starts** | 1001, Charmander, 100, 10, Ember, 1, Scratch, 1 |

* **Simple Value**: The header occupies a single column (e.g., Id, Name).  
* **Nested Object**: The header is a merged cell, with its properties defined below (e.g., BaseProperty \-\> Hp, Attack).  
* **List of Objects**: The header is a merged cell, with a repeating object structure below (e.g., Skills \-\> Name, Level, Name, Level...).  
* **List of Primitives**: The header is a merged cell, with repeating, same-named simple properties below (e.g., LiveScene \-\> Scene, Scene...).

#### **4\. Running the Tool**

Place your Excel files (with sheet names ending in \_Table, e.g., Boss\_Table.xlsx) into the InputPath directory specified in AppConfig.xml. Then, run the tool:

dotnet run

Or simply execute the compiled .exe. The tool will generate DLL, Schemas, and Json folders in your OutputPath.

#### **5\. Using the Generated DLL in Your Project**

1. **Reference the DLL**: Copy MyGame.Configs.dll and MyGame.Configs.xml from the output/DLL/ directory into your main project and add a reference to the DLL.  
2. **Write Loading Code**: In your main project, use the DeserializerFactory helper class (which can be copied from the tool's exception message or pre-added to your project) to load the data.  
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


## **Project Dependencies**

* [NPOI](https://www.google.com/search?q=https://github.com/nissim-aj/npoi) \- For reading Excel files.  
* [Newtonsoft.Json](https://www.newtonsoft.com/json) \- For data serialization and the deserializer helper class.  
* [Microsoft.CodeAnalysis.CSharp](https://www.nuget.org/packages/Microsoft.CodeAnalysis.CSharp/) (Roslyn) \- For dynamic C\# code compilation.

## **License**

This project is licensed under the [MIT](https://opensource.org/licenses/MIT) License.