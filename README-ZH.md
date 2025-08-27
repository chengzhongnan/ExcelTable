# **Excel ���ñ��빤�� (Excel to DLL Configuration Compiler)**

## **��Ŀ���**

����Ŀ��һ��ǿ��� .NET �����й��ߣ�ּ�ڽ����ӵ� Excel �����ļ���.xlsx/.xls��ֱ�ӱ����һ�����Ͱ�ȫ������ʹ�á��������� **.NET Standard 2.1 DLL**��

����Ϸ��������ҪƵ���������õ�Ӧ�ó����У���ͳ���ֶ�����ճ�����д��������ķ�ʽЧ�ʵ��������׳��������߽����������Զ������ÿ����߿���רע��ҵ���߼���ͬʱΪ�߻�����Ӫ��Ա�ṩ������Ϥ�� Excel ��Ϊ���ù��ߡ�

## **��������**

* **��̬�����༶��ͷ**����׼ȷʶ���ӵ� Excel ��ͷ�ṹ�������ϲ���Ԫ������Ƕ�׶�����б�  
* **�Զ����� C\# ��**�����ݱ�ͷ�ṹ�Զ����ɶ�Ӧ��ǿ���� C\# �����ࡣ  
* **���Ͱ�ȫ**�����ɵ� DataManager �ṩ���Ͱ�ȫ�����ݷ��ʽӿڣ��ž�������д���ֶ���������ת�������µ� Bug��  
* **����Ϊ DLL**�������ɵ� C\# ��� JSON ����ֱ�ӱ����һ�������� .NET Standard 2.1 DLL���������κ� .NET ��Ŀ�����á�  
* **Ƕ�� JSON ����**���ڱ���ʱ���� Excel �е�����ת��Ϊ JSON ����Ϊ��Ƕ��Դ����� DLL��ʵ�����ݵ��԰�����  
* **���� XML �ļ�ע��**���� Excel ��ͷ�е�ע���б���Ϊ XML �ļ�ע�ͣ��� Visual Studio ���ṩ������ IntelliSense ������ʾ��  
* **�������ݼ���**�����ɵ� DataManager ֧�ִ� DLL ��Ƕ��Դ���ⲿ�ļ�������ģʽ�������ã����㿪�����ԺͲ���  
* **����ע�����**��DataManager ������ JSON �������������û�������Ŀ���ṩ�Լ��Ľ����ص�������  
* **�ⲿ������**�����߱��������·�������þ��� AppConfig.xml ���ƣ������޸Ĵ��롣

## **���ʹ��**

#### **1\. ����Ҫ��**

* ��װ [.NET SDK](https://dotnet.microsoft.com/download) (�Ƽ� .NET 8.0 ����߰汾)��

#### **2\. ���� AppConfig.xml**

�ڹ��ߵ�ִ��Ŀ¼�£����������� AppConfig.xml��

\<?xml version="1.0" encoding="utf-8" ?\>  
\<Configuration\>  
  \<\!-- �Ƽ�ʹ������ڴ˹��� .exe �����·�� \--\>  
  \<Paths\>  
    \<\!-- �������� .xlsx �ļ��������ļ��� \--\>  
    \<InputPath\>xlsx\</InputPath\>  
    \<\!-- ���������ļ��ĸ�����ļ��� \--\>  
    \<OutputPath\>output\</OutputPath\>  
  \</Paths\>  
  \<Settings\>  
    \<\!-- ���ɵ� DLL �ĳ������ƣ�Ҳ���� C\# ��ĸ������ռ� \--\>  
    \<AssemblyName\>MyGame.Configs\</AssemblyName\>  
  \</Settings\>  
\</Configuration\>

#### **3\. Excel ���淶**

Ϊ���ù�������ȷ���������� Excel �����Ҫ��ѭ����**���б�ͷ**�ṹ��

| �� N �� | Ŀ�� | ʾ�� |
| :---- | :---- | :---- |
| **�� 1 ��** | **�ṹ���� (��)** | Id, Name, BaseProperty (�ϲ���Ԫ��), Skills (�ϲ���Ԫ��) |
| **�� 2 ��** | **�ṹ���� (��)** | (Id��Ϊ��), (Name��Ϊ��), Hp, Attack, Name, Level, Name, Level |
| **�� 3 ��** | **���Ͷ���** | int, string, int, int, string, int, string, int |
| **�� 4 ��** | **ע��/˵��** | ΨһID, ����, ����ֵ, ������, ��������, ���ܵȼ�, ��������, ���ܵȼ� |
| **�� 5 ��** | **���ݿ�ʼ** | 1001, С����, 100, 10, ��, 1, ץ, 1 |

* **��һֵ**����ͷֻռһ�� (�� Id, Name)��  
* **Ƕ�׶���**����ͷΪ�ϲ���Ԫ�����·������˶���ĸ������� (�� BaseProperty \-\> Hp, Attack)��  
* **�����б�**����ͷΪ�ϲ���Ԫ�����·����ظ��Ķ���ṹ (�� Skills \-\> Name, Level, Name, Level...)��  
* **�������б�**����ͷΪ�ϲ���Ԫ�����·����ظ���ͬ�������� (�� LiveScene \-\> Scene, Scene...)��

#### **4\. ���й���**

������ Excel �ļ���Sheet ҳ�������� \_Table ��β���� Boss\_Table.xlsx������ AppConfig.xml ��ָ���� InputPath �ļ����ڡ�  
Ȼ��ִ�й��ߣ�  
dotnet run

��ֱ��ִ�б���õ� .exe �ļ���

���߽��� OutputPath ������ DLL��Schemas �� Json �ļ��С�

#### **5\. ����Ŀ��ʹ�����ɵ� DLL**

1. **���� DLL**���� output/DLL/ Ŀ¼�µ� MyGame.Configs.dll �� MyGame.Configs.xml ���Ƶ���������Ŀ�У�����Ӷ� DLL �����á�  
2. **��д���ش���**������������Ŀ�У�ʹ�� DeserializerFactory �����ࣨ�ɴӹ��ߵ��쳣��Ϣ�и��ƣ���Ԥ�ȼ�����Ŀ�����������ݡ�  
   using MyGame.Configs; // �������ɵ������ռ�  
   using System.Reflection;

   public class GameInitializer  
   {  
       public void Initialize()  
       {  
           // ��ȡ���� DLL �ĳ���  
           var configAssembly \= typeof(DataManager).Assembly;

           // ʹ�ù��������Զ��������б��� JSON ������  
           var deserializers \= DeserializerFactory.CreateJsonDeserializersForAllTables(configAssembly, "MyGame.Configs");

           // ��ʽһ���� DLL ��Ƕ��Դ���� (������ʽ����)  
           DataManager.LoadAllFromEmbedded(configAssembly, deserializers);

           // ��ʽ�������ⲿ�ļ��м��� (���ڿ�������)  
           // string configPath \= "path/to/your/json/files";  
           // DataManager.LoadAllFromDirectory(configPath, deserializers, ".json");

           // \--- ������ɣ����԰�ȫʹ���� \---  
           var boss \= DataManager.GetBossById(1001);  
           if (boss \!= null)  
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

## **��Ŀ����**

* [NPOI](https://www.google.com/search?q=https://github.com/nissim-aj/npoi) \- ���ڶ�ȡ Excel �ļ���  
* [Newtonsoft.Json](https://www.newtonsoft.com/json) \- �����������л��͸����ࡣ  
* [Microsoft.CodeAnalysis.CSharp](https://www.nuget.org/packages/Microsoft.CodeAnalysis.CSharp/) (Roslyn) \- ���ڶ�̬���� C\# ���롣

## **���֤**

����Ŀ���� [MIT](https://opensource.org/licenses/MIT) ���֤��